//! Access log for `vybex --serve`. Phase 1: text-format to stderr.

use std::net::SocketAddr;
use std::time::Duration;

pub fn log_request(
    remote: SocketAddr,
    method: &str,
    path: &str,
    protocol: &str,
    status: u16,
    latency: Duration,
) {
    let ts = timestamp();
    eprintln!(
        "[{ts}] {remote} {method} {path} {protocol} {status} {}ms",
        latency.as_millis()
    );
}

fn timestamp() -> String {
    // Minimal Phase 1: a short local timestamp. Phase 2+ uses RFC 3339
    // via `httpdate` or a `chrono` format.
    use std::time::{SystemTime, UNIX_EPOCH};
    let secs = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_secs())
        .unwrap_or(0);
    format!("{secs}")
}
