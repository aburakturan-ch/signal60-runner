const SITE_URL = (
  process.env.SIGNAL60_SITE_URL ||
  "https://binance-hourly-heatmap.burakturan000.chatgpt.site"
).replace(/\/+$/, "");
const RUNNER_SECRET = process.env.SIGNAL60_RUNNER_SECRET;
const RUNNER_ID =
  process.env.SIGNAL60_RUNNER_ID || "github-actions-v1";

const BINANCE_BASES = process.env.SIGNAL60_BINANCE_BASES
  ? process.env.SIGNAL60_BINANCE_BASES.split(",")
      .map((value) => value.trim().replace(/\/+$/, ""))
      .filter(Boolean)
  : [
      "https://data-api.binance.vision",
      "https://api.binance.com",
      "https://api-gcp.binance.com",
      "https://api1.binance.com",
      "https://api2.binance.com",
      "https://api3.binance.com",
      "https://api4.binance.com",
    ];
const KLINE_BATCH_SIZE = 6;
const FUSION_RISK_KLINE_BATCH_SIZE = 12;
const INDEX_BATCH_SIZE = 8;
const MINIMUM_DEEP_MODELS = 60;
const BINANCE_MAX_ATTEMPTS = 3;

let githubOidcToken;
const cycleFetchCache = new Map();

async function getAuthorization() {
  if (RUNNER_SECRET && RUNNER_SECRET.length >= 32) {
    return `Bearer ${RUNNER_SECRET}`;
  }
  if (githubOidcToken) return `Bearer ${githubOidcToken}`;

  const requestUrl = process.env.ACTIONS_ID_TOKEN_REQUEST_URL;
  const requestToken = process.env.ACTIONS_ID_TOKEN_REQUEST_TOKEN;
  if (!requestUrl || !requestToken) {
    throw new Error(
      "GitHub Actions OIDC is unavailable and no protected runner secret is configured",
    );
  }
  const tokenUrl = new URL(requestUrl);
  tokenUrl.searchParams.set("audience", `${SITE_URL}/api/market`);
  const response = await fetch(tokenUrl, {
    headers: {
      Accept: "application/json",
      Authorization: `Bearer ${requestToken}`,
    },
  });
  if (!response.ok) {
    throw new Error(`GitHub Actions identity request failed (${response.status})`);
  }
  const body = await response.json();
  if (typeof body.value !== "string" || body.value.length < 100) {
    throw new Error("GitHub Actions returned an invalid identity token");
  }
  githubOidcToken = body.value;
  return `Bearer ${githubOidcToken}`;
}

async function fetchWithTimeout(url, init = {}, timeoutMs = 20_000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(url, {
      ...init,
      signal: controller.signal,
      headers: {
        Accept: "application/json",
        "User-Agent": "SIGNAL60-External-Runner/1.0",
        ...init.headers,
      },
    });
  } finally {
    clearTimeout(timer);
  }
}

async function fetchBinanceUncached(path) {
  const failures = [];
  for (let attempt = 0; attempt < BINANCE_MAX_ATTEMPTS; attempt += 1) {
    let retryAfterMs = 0;
    for (const base of BINANCE_BASES) {
      try {
        const response = await fetchWithTimeout(`${base}${path}`);
        if (!response.ok) {
          failures.push(`${new URL(base).hostname}:${response.status}`);
          if (response.status === 429) {
            const retryAfter = Number(response.headers.get("retry-after"));
            retryAfterMs = Math.max(
              retryAfterMs,
              Number.isFinite(retryAfter) ? retryAfter * 1_000 : 0,
            );
          }
          if (response.status >= 500 || response.status === 429) continue;
          continue;
        }
        return await response.json();
      } catch (error) {
        failures.push(
          `${new URL(base).hostname}:${
            error instanceof Error ? error.name : "fetch_error"
          }`,
        );
      }
    }
    if (attempt < BINANCE_MAX_ATTEMPTS - 1) {
      const exponentialMs = 500 * 2 ** attempt;
      await new Promise((resolve) =>
        setTimeout(resolve, Math.max(exponentialMs, retryAfterMs)),
      );
    }
  }
  throw new Error(
    `Official Binance endpoints unavailable (${failures.slice(-8).join(", ")})`,
  );
}

async function fetchBinance(path) {
  if (!cycleFetchCache.has(path)) {
    cycleFetchCache.set(path, fetchBinanceUncached(path));
  }
  return cycleFetchCache.get(path);
}

async function postRunner(payload, attempts = 2) {
  let lastError;
  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    try {
      const response = await fetchWithTimeout(
        `${SITE_URL}/api/market`,
        {
          method: "POST",
          headers: {
            Authorization: await getAuthorization(),
            "Content-Type": "application/json",
          },
          body: JSON.stringify(payload),
        },
        110_000,
      );
      const body = await response.json().catch(() => ({}));
      if (!response.ok) {
        const detail =
          typeof body.detail === "string"
            ? body.detail
            : typeof body.error === "string"
              ? body.error
              : `HTTP ${response.status}`;
        throw new Error(detail);
      }
      return body;
    } catch (error) {
      lastError = error;
      if (attempt < attempts) {
        await new Promise((resolve) => setTimeout(resolve, 1_500 * attempt));
      }
    }
  }
  throw lastError;
}

function compactTickers(rows, bookRows) {
  const books = new Map(
    bookRows.flatMap((row) =>
      row &&
      typeof row === "object" &&
      typeof row.symbol === "string"
        ? [[row.symbol, row]]
        : [],
    ),
  );
  return rows.flatMap((row) => {
    if (
      !row ||
      typeof row !== "object" ||
      typeof row.symbol !== "string" ||
      !row.symbol.endsWith("USDT")
    ) {
      return [];
    }
    const book = books.get(row.symbol);
    return [
      {
        symbol: row.symbol,
        lastPrice: String(row.lastPrice ?? ""),
        priceChangePercent: String(row.priceChangePercent ?? ""),
        quoteVolume: String(row.quoteVolume ?? ""),
        count: Number(row.count ?? 0),
        highPrice: String(row.highPrice ?? ""),
        lowPrice: String(row.lowPrice ?? ""),
        bidPrice: String(book?.bidPrice ?? ""),
        askPrice: String(book?.askPrice ?? ""),
      },
    ];
  });
}

function compactKlines(rows) {
  return rows.flatMap((row) => {
    if (!Array.isArray(row) || row.length < 8) return [];
    return [
      [
        Number(row[0]),
        Number(row[4]),
        Number(row[6]),
        Number(row[7]),
      ],
    ];
  });
}

function klineVwap(row) {
  if (!Array.isArray(row) || row.length < 8) return null;
  const close = Number(row[4]);
  const baseVolume = Number(row[5]);
  const quoteVolume = Number(row[7]);
  if (
    Number.isFinite(baseVolume) &&
    baseVolume > 0 &&
    Number.isFinite(quoteVolume) &&
    quoteVolume > 0
  ) {
    return quoteVolume / baseVolume;
  }
  return Number.isFinite(close) && close > 0 ? close : null;
}

async function fetchSyntheticIndexSubmission(plan) {
  if (!plan || plan.required !== true) return null;
  if (
    !plan.epoch ||
    typeof plan.epoch.id !== "string" ||
    !Number.isFinite(Number(plan.epoch.startAt)) ||
    !Number.isFinite(Number(plan.epoch.endAt)) ||
    !Array.isArray(plan.pairs)
  ) {
    throw new Error("SIGNAL/60 returned an invalid synthetic-index plan");
  }

  const startAt = Number(plan.epoch.startAt);
  const endAt = Number(plan.epoch.endAt);
  if (endAt - startAt !== 24 * 60 * 1_000) {
    throw new Error("SIGNAL/60 synthetic-index epoch is not 24 minutes");
  }

  const observations = [];
  const failedPairs = [];
  for (let start = 0; start < plan.pairs.length; start += INDEX_BATCH_SIZE) {
    const batch = plan.pairs.slice(start, start + INDEX_BATCH_SIZE);
    const results = await Promise.allSettled(
      batch.map(async (pair) => {
        if (typeof pair !== "string" || !/^[A-Z0-9]+USDT$/.test(pair)) {
          throw new Error("invalid index pair");
        }
        const rows = await fetchBinance(
          `/api/v3/klines?symbol=${encodeURIComponent(
            pair,
          )}&interval=1m&startTime=${startAt}&endTime=${
            endAt - 1
          }&limit=24`,
        );
        if (!Array.isArray(rows) || rows.length < 20 || rows.length > 24) {
          throw new Error(`${pair} returned ${rows?.length ?? 0} index bars`);
        }
        const startPrice = klineVwap(rows[0]);
        const endPrice = klineVwap(rows.at(-1));
        if (!startPrice || !endPrice) {
          throw new Error(`${pair} index VWAP is invalid`);
        }
        observations.push({
          pair,
          startPrice,
          endPrice,
          sampleCount: rows.length,
        });
      }),
    );
    results.forEach((result, index) => {
      if (result.status === "rejected") {
        failedPairs.push(batch[index] ?? "unknown");
      }
    });
  }
  return {
    epoch: {
      id: plan.epoch.id,
      startAt,
      endAt,
    },
    observations,
    failedPairs,
  };
}

async function runCycle() {
  const startedAt = Date.now();
  const cycleStartedAt =
    Math.floor(startedAt / (5 * 60 * 1_000)) * (5 * 60 * 1_000);
  const [rawTickers, rawBooks] = await Promise.all([
    fetchBinance("/api/v3/ticker/24hr"),
    fetchBinance("/api/v3/ticker/bookTicker"),
  ]);
  if (!Array.isArray(rawTickers) || !Array.isArray(rawBooks)) {
    throw new Error("Binance ticker or order-book response is not an array");
  }
  const tickers = compactTickers(rawTickers, rawBooks);
  const plan = await postRunner({
    action: "plan",
    runnerId: RUNNER_ID,
    cycleStartedAt,
    tickers,
  });
  if (
    !Array.isArray(plan.pairs) ||
    plan.pairs.length < MINIMUM_DEEP_MODELS
  ) {
    throw new Error("SIGNAL/60 returned an incomplete deep-model plan");
  }
  if (
    !Array.isArray(plan.fusionRiskPairs) ||
    plan.fusionRiskPairs.length < plan.pairs.length
  ) {
    throw new Error("SIGNAL/60 returned an incomplete Fusion Risk universe");
  }
  if (
    !plan.syntheticIndex ||
    typeof plan.syntheticIndex.required !== "boolean" ||
    !plan.syntheticIndex.epoch ||
    typeof plan.syntheticIndex.epoch.id !== "string"
  ) {
    throw new Error("SIGNAL/60 returned an incomplete synthetic-index plan");
  }
  let syntheticIndexFailure = null;
  const syntheticIndexPromise = fetchSyntheticIndexSubmission(
    plan.syntheticIndex,
  ).catch((error) => {
    syntheticIndexFailure =
      error instanceof Error ? error.message : "synthetic index fetch failed";
    return null;
  });

  const datasets = {};
  const failures = [];
  for (let start = 0; start < plan.pairs.length; start += KLINE_BATCH_SIZE) {
    const batch = plan.pairs.slice(start, start + KLINE_BATCH_SIZE);
    const results = await Promise.allSettled(
      batch.map(async ({ pair }) => {
        if (typeof pair !== "string" || !/^[A-Z0-9]+USDT$/.test(pair)) {
          throw new Error("invalid planned pair");
        }
        const rows = await fetchBinance(
          `/api/v3/klines?symbol=${encodeURIComponent(
            pair,
          )}&interval=5m&limit=1000`,
        );
        if (!Array.isArray(rows)) {
          throw new Error(`${pair} kline response is not an array`);
        }
        datasets[pair] = compactKlines(rows);
      }),
    );
    results.forEach((result, index) => {
      if (result.status === "rejected") {
        failures.push(
          `${batch[index]?.pair ?? "unknown"}:${
            result.reason instanceof Error
              ? result.reason.message
              : "fetch failed"
          }`,
        );
      }
    });
  }
  const deepPairs = new Set(
    plan.pairs.flatMap(({ pair }) =>
      typeof pair === "string" ? [pair] : [],
    ),
  );
  const fusionRiskOnlyPairs = plan.fusionRiskPairs.filter(
    ({ pair }) => typeof pair === "string" && !deepPairs.has(pair),
  );
  for (
    let start = 0;
    start < fusionRiskOnlyPairs.length;
    start += FUSION_RISK_KLINE_BATCH_SIZE
  ) {
    const batch = fusionRiskOnlyPairs.slice(
      start,
      start + FUSION_RISK_KLINE_BATCH_SIZE,
    );
    const results = await Promise.allSettled(
      batch.map(async ({ pair }) => {
        if (typeof pair !== "string" || !/^[A-Z0-9]+USDT$/.test(pair)) {
          throw new Error("invalid Fusion Risk pair");
        }
        const rows = await fetchBinance(
          `/api/v3/klines?symbol=${encodeURIComponent(
            pair,
          )}&interval=5m&limit=300`,
        );
        if (!Array.isArray(rows)) {
          throw new Error(`${pair} Fusion Risk kline response is not an array`);
        }
        datasets[pair] = compactKlines(rows);
      }),
    );
    results.forEach((result, index) => {
      if (result.status === "rejected") {
        failures.push(
          `${batch[index]?.pair ?? "unknown"}:${
            result.reason instanceof Error
              ? result.reason.message
              : "Fusion Risk fetch failed"
          }`,
        );
      }
    });
  }
  if (Object.keys(datasets).length < MINIMUM_DEEP_MODELS) {
    throw new Error(
      `Only ${Object.keys(datasets).length} deep datasets loaded (${failures
        .slice(0, 3)
        .join(", ")})`,
    );
  }
  const syntheticIndex = await syntheticIndexPromise;

  const result = await postRunner(
    {
      action: "execute",
      runnerId: RUNNER_ID,
      cycleStartedAt,
      tickers,
      datasets,
      failedSymbols: failures.map((failure) => failure.split(":")[0]),
      ...(syntheticIndex ? { syntheticIndex } : {}),
    },
    2,
  );
  const summary = {
    ok: true,
    generatedAt: result.generatedAt,
    status: result.status,
    scanned: result.universe?.scanned ?? null,
    modeled: result.universe?.deeplyModeled ?? null,
    failed: result.universe?.failed ?? failures.length,
    paperCycleExecuted: result.trading?.cycle?.executed ?? null,
    duplicateCycle: result.trading?.cycle?.duplicate ?? false,
    ordersCreated: result.trading?.cycle?.ordersCreated ?? null,
    positions: result.trading?.positions?.length ?? null,
    orders: result.trading?.orders?.length ?? null,
    fusionRiskTracked: result.fusionRisk?.universe?.tracked ?? null,
    fusionRiskModeled: result.fusionRisk?.universe?.modeled ?? null,
    fusionRiskFailed: result.fusionRisk?.universe?.failed ?? null,
    fusionRiskOrdersCreated:
      result.fusionRisk?.portfolio?.cycle?.ordersCreated ?? null,
    fusionRiskPositions:
      result.fusionRisk?.portfolio?.positions?.length ?? null,
    s60PlanRequired: plan.syntheticIndex.required,
    s60PlannedEpoch: plan.syntheticIndex.epoch.id,
    s60Observed: syntheticIndex?.observations?.length ?? 0,
    s60FailedPairs: syntheticIndex?.failedPairs?.length ?? 0,
    s60Epoch: result.synthetic?.epoch?.lastSettledId ?? null,
    s60SettlementExecuted:
      result.synthetic?.lastSettlement?.id === syntheticIndex?.epoch?.id,
    s60Included: result.synthetic?.oracle?.included ?? null,
    s60FetchFailure: syntheticIndexFailure,
    warnings: Array.isArray(result.warnings) ? result.warnings.length : 0,
    elapsedMs: Date.now() - startedAt,
  };
  process.stdout.write(`${JSON.stringify(summary)}\n`);
}

runCycle().catch((error) => {
  const safeMessage =
    error instanceof Error ? error.message : "Unknown runner failure";
  process.stderr.write(`SIGNAL60 runner failed: ${safeMessage}\n`);
  process.exitCode = 1;
});
