// node:http — Node-style HTTP module for Vybe.
//
// Phase 6 (re-export-only): exposes the programmatic server primitives
// from `vybe:http/server` under Node's specifier. Full Node-API-shape
// translation (`http.createServer((req, res) => ...)` with req/res
// objects) is a follow-up phase that requires local function bodies
// in the adapter — not just re-exports.
//
// Layer 3 rule: Node's idiom lives in JS, not Rust. This file IS that
// layer.

export { listen } from "node:http";
