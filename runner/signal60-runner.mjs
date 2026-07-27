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
const KLINE_BATCH_SIZE = 4;

let githubOidcToken;

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

async function fetchBinance(path) {
  const failures = [];
  for (const base of BINANCE_BASES) {
    try {
      const response = await fetchWithTimeout(`${base}${path}`);
      if (!response.ok) {
        failures.push(`${new URL(base).hostname}:${response.status}`);
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
  throw new Error(`Official Binance endpoints unavailable (${failures.join(", ")})`);
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

function compactTickers(rows) {
  return rows.flatMap((row) => {
    if (
      !row ||
      typeof row !== "object" ||
      typeof row.symbol !== "string" ||
      !row.symbol.endsWith("USDT")
    ) {
      return [];
    }
    return [
      {
        symbol: row.symbol,
        lastPrice: String(row.lastPrice ?? ""),
        priceChangePercent: String(row.priceChangePercent ?? ""),
        quoteVolume: String(row.quoteVolume ?? ""),
        count: Number(row.count ?? 0),
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

async function runCycle() {
  const cycleStartedAt = Date.now();
  const rawTickers = await fetchBinance("/api/v3/ticker/24hr");
  if (!Array.isArray(rawTickers)) {
    throw new Error("Binance ticker response is not an array");
  }
  const tickers = compactTickers(rawTickers);
  const plan = await postRunner({
    action: "plan",
    runnerId: RUNNER_ID,
    cycleStartedAt,
    tickers,
  });
  if (!Array.isArray(plan.pairs) || plan.pairs.length < 8) {
    throw new Error("SIGNAL/60 returned an incomplete deep-model plan");
  }

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
  if (Object.keys(datasets).length < 8) {
    throw new Error(
      `Only ${Object.keys(datasets).length} deep datasets loaded (${failures
        .slice(0, 3)
        .join(", ")})`,
    );
  }

  const result = await postRunner(
    {
      action: "execute",
      runnerId: RUNNER_ID,
      cycleStartedAt,
      tickers,
      datasets,
    },
    2,
  );
  const summary = {
    ok: true,
    generatedAt: result.generatedAt,
    status: result.status,
    scanned: result.universe?.scanned ?? null,
    modeled: result.universe?.deeplyModeled ?? null,
    paperCycleExecuted: result.trading?.cycle?.executed ?? null,
    positions: result.trading?.positions?.length ?? null,
    orders: result.trading?.orders?.length ?? null,
    warnings: Array.isArray(result.warnings) ? result.warnings.length : 0,
    elapsedMs: Date.now() - cycleStartedAt,
  };
  process.stdout.write(`${JSON.stringify(summary)}\n`);
}

runCycle().catch((error) => {
  const safeMessage =
    error instanceof Error ? error.message : "Unknown runner failure";
  process.stderr.write(`SIGNAL60 runner failed: ${safeMessage}\n`);
  process.exitCode = 1;
});
