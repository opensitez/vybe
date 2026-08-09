//! Configuration for `vybex --serve`. Parsed from CLI flags in `main.rs`.

use std::path::PathBuf;

#[derive(Debug, Clone)]
pub struct ServeConfig {
    /// host:port to bind on.
    pub bind: String,
    /// Root directory to serve.
    pub root: PathBuf,
    /// If true, scripts run with `Capabilities::all()`. Otherwise
    /// they use the restricted sandbox capability set.
    pub no_sandbox: bool,
    /// Maximum request body size in bytes.
    pub max_body: usize,
    /// Maximum total header size in bytes.
    pub max_header: usize,
    /// Per-request total timeout (header read + body read + response write).
    /// Phase 2 will split this into separate limits.
    pub timeout_secs: u64,
    /// Filenames to try in order when a request resolves to a directory.
    pub index_files: Vec<String>,
    /// Shutdown notifier set by `run()`. In-flight request handlers race
    /// their per-request timeout against this; on Ctrl+C we flip it and
    /// hung scripts release with a 503 instead of blocking the drain.
    /// Not serialisable — `ServeConfig::default()` leaves it `None` so
    /// tests and programmatic callers work without wiring one up.
    #[allow(dead_code)]
    pub shutdown: Option<std::sync::Arc<tokio::sync::Notify>>,
    /// `--cold`: build a whole VM per request instead of resetting a warm one.
    /// The control the warm pool is measured and isolation-diffed against, and
    /// the escape hatch if a leak between requests ever turns up in the field.
    pub cold: bool,
    /// `--pool N`: warm VM threads. Each holds one booted VM for the life of
    /// the process, so this is also the ceiling on concurrent script requests
    /// — the rest queue. 0 means "one per core".
    pub pool: usize,
    /// `--no-cache`: recompile every request even when nothing changed.
    /// The cache validates against every file a compile read, so this should
    /// only be needed for an editor that rewrites a file to the same length
    /// within one mtime tick.
    pub no_cache: bool,
}

impl Default for ServeConfig {
    fn default() -> Self {
        Self {
            bind: "127.0.0.1:8080".into(),
            root: PathBuf::from("."),
            no_sandbox: true,
            max_body: 10 * 1024 * 1024,
            max_header: 16 * 1024,
            timeout_secs: 30,
            index_files: vec![
                "index.php".into(),
                "index.html".into(),
                "index.htm".into(),
                "index.js".into(),
                "index.py".into(),
                "index.rb".into(),
            ],
            shutdown: None,
            cold: false,
            pool: 0,
            no_cache: false,
        }
    }
}
