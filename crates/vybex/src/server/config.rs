//! Configuration for `vybex --serve`. Parsed from CLI flags in `main.rs`.

use std::path::PathBuf;

#[derive(Debug, Clone)]
pub struct ServeConfig {
    /// host:port to bind on.
    pub bind: String,
    /// Root directory to serve.
    pub root: PathBuf,
    /// If true, scripts run with `Capabilities::all()`. Otherwise
    /// `Capabilities::safe()` plus `HttpServer` for `vybe:http` access.
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
}

impl Default for ServeConfig {
    fn default() -> Self {
        Self {
            bind: "127.0.0.1:8080".into(),
            root: PathBuf::from("."),
            no_sandbox: false,
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
        }
    }
}
