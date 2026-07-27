# SIGNAL/60 external runner

This runner is the only process allowed to publish market cycles. The public
dashboard performs read-only `GET /api/market` requests and therefore cannot
create signals, orders, forecasts, or paper-ledger mutations.

The provided GitHub Actions workflow runs every five minutes and can also be
started manually. It:

1. Downloads the full official Binance Spot 24-hour ticker universe.
2. Requests an authenticated deep-model plan from SIGNAL/60.
3. Downloads 1,000 closed 5-minute bars for the selected markets.
4. Sends the compact datasets to the protected server model.
5. Publishes the resulting snapshot only after forecast verification and the
   idempotent paper-trading cycle complete.

No long-lived repository secret is required. The workflow requests a
short-lived GitHub Actions OIDC identity token, and the site accepts it only
when it was issued for `aburakturan-ch/signal60-runner`, the `main` branch, and
a push, scheduled, or manual workflow event. GitHub schedule delivery can
occasionally be delayed; the dashboard marks the runner stale after twelve
minutes and never executes a trade from a public read.
