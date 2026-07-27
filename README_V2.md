# SIGNAL/60 V2

Auditable Binance Spot universe scanner.

- Broad layer: every currently trading Spot symbol returned by Binance exchangeInfo.
- Live layer: Binance all-market mini-ticker WebSocket updates all changed symbols once per second.
- Deep layer: highest-ranked liquid candidates receive 5m/15m/1h/4h candle analysis; top subset also receives 100-level order-book analysis.
- No trained-AI claim and no live-order code.

Validation: `node --test tests/engine.test.mjs`
