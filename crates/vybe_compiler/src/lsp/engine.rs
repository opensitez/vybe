//! Background analysis engine — runs parsers on a dedicated thread.
//!
//! Design:
//! - NEVER blocks the UI thread. All communication via channels.
//! - Debounced: rapid edits coalesce; only the latest version is parsed.
//! - Cancellation: stale requests are dropped before parsing starts.
//! - Version-tracked: results carry a version; the UI discards stale results.
//! - On initial file open, the editor renders immediately; analysis arrives later.

use super::extract;
use super::symbols::*;
use crate::ast::Lang;
use crossbeam_channel::{Receiver, Sender, TryRecvError};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};
use std::thread;
use std::time::{Duration, Instant};

/// Request from the UI thread to the analysis thread.
pub enum AnalysisRequest {
    /// File opened or content changed. Engine will debounce and parse.
    Update {
        uri: String,
        content: String,
        version: u64,
    },
    /// Shutdown the analysis thread.
    Shutdown,
}

/// Event from the analysis thread back to the UI.
pub enum AnalysisEvent {
    /// Full analysis result (symbols + diagnostics) for a file.
    Analysis(AnalysisResult),
}

/// The analysis engine. Create one, send requests, poll for events.
pub struct AnalysisEngine {
    tx: Sender<AnalysisRequest>,
    rx: Receiver<AnalysisEvent>,
    version: Arc<AtomicU64>,
}

const DEBOUNCE_MS: u64 = 250;

impl AnalysisEngine {
    pub fn new() -> Self {
        let (req_tx, req_rx) = crossbeam_channel::unbounded::<AnalysisRequest>();
        let (evt_tx, evt_rx) = crossbeam_channel::unbounded::<AnalysisEvent>();
        let version = Arc::new(AtomicU64::new(0));
        let ver_clone = version.clone();

        thread::Builder::new()
            .name("vybe-lsp-analysis".into())
            .spawn(move || {
                analysis_loop(req_rx, evt_tx, ver_clone);
            })
            .expect("failed to spawn analysis thread");

        Self {
            tx: req_tx,
            rx: evt_rx,
            version,
        }
    }

    /// Send an update with auto-incrementing version. Returns the version number.
    /// Never blocks.
    pub fn update(&self, uri: String, content: String) -> u64 {
        let v = self.version.fetch_add(1, Ordering::Relaxed) + 1;
        self.tx
            .send(AnalysisRequest::Update {
                uri,
                content,
                version: v,
            })
            .ok();
        v
    }

    /// Poll for events. Returns None if no events are ready. Never blocks.
    pub fn try_recv(&self) -> Option<AnalysisEvent> {
        match self.rx.try_recv() {
            Ok(event) => Some(event),
            Err(TryRecvError::Empty) => None,
            Err(TryRecvError::Disconnected) => None,
        }
    }
}

impl Drop for AnalysisEngine {
    fn drop(&mut self) {
        self.tx.send(AnalysisRequest::Shutdown).ok();
    }
}

/// Synchronous analysis — parse + extract symbols + diagnostics.
/// Used by code_editor for inline analysis when no external LSP is running.
pub fn analyze(uri: &str, content: &str) -> AnalysisResult {
    let lang = extract::detect_language(uri);
    let keywords = extract::language_keywords(lang);
    let (symbols, diagnostics) = parse_content(lang, content);
    AnalysisResult {
        uri: uri.to_string(),
        version: 0,
        symbols,
        diagnostics,
        keywords,
    }
}

fn parse_content(lang: Lang, content: &str) -> (Vec<Symbol>, Vec<LspDiagnostic>) {
    crate::ensure_languages_registered();
    let parse_fn: Option<fn(&str) -> Result<crate::ast::Module, String>> = match lang {
        Lang::VB => vybe_bytecode::registry::find("vb").map(|p| p.parse),
        Lang::JavaScript => vybe_bytecode::registry::find("js").map(|p| p.parse),
        Lang::CSharp => vybe_bytecode::registry::find("csharp").map(|p| p.parse),
        Lang::Python => vybe_bytecode::registry::find("python").map(|p| p.parse),
        Lang::Ruby => vybe_bytecode::registry::find("ruby").map(|p| p.parse),
        Lang::PHP => vybe_bytecode::registry::find("php").map(|p| p.parse),
        Lang::Dart => vybe_bytecode::registry::find("dart").map(|p| p.parse),
        Lang::Pascal => vybe_bytecode::registry::find("pascal").map(|p| p.parse),
        Lang::Cobol => vybe_bytecode::registry::find("cobol").map(|p| p.parse),
        _ => None,
    };

    if let Some(parser) = parse_fn {
        match parser(content) {
            Ok(module) => (extract::extract_symbols(&module), Vec::new()),
            Err(msg) => {
                let (line, col) = parse_error_location(&msg);
                (
                    Vec::new(),
                    vec![LspDiagnostic {
                        line,
                        col,
                        end_col: col + 10,
                        message: msg,
                        severity: DiagSeverity::Error,
                    }],
                )
            }
        }
    } else {
        (Vec::new(), Vec::new())
    }
}

fn parse_error_location(msg: &str) -> (u32, u32) {
    // Try pest format: " --> LINE:COL"
    if let Some(arrow_pos) = msg.find(" --> ") {
        let after = &msg[arrow_pos + 4..];
        let parts: Vec<&str> = after.splitn(2, ':').collect();
        if parts.len() == 2 {
            if let Ok(line) = parts[0].trim().parse::<u32>() {
                let col = parts[1]
                    .trim()
                    .split(|c: char| !c.is_ascii_digit())
                    .next()
                    .and_then(|s| s.parse::<u32>().ok())
                    .unwrap_or(0);
                return (line.saturating_sub(1), col.saturating_sub(1));
            }
        }
    }
    // Try "line N" format
    if let Some(idx) = msg.to_lowercase().find("line ") {
        let after = &msg[idx + 5..];
        if let Some(line) = after
            .split(|c: char| !c.is_ascii_digit())
            .next()
            .and_then(|s| s.parse::<u32>().ok())
        {
            return (line.saturating_sub(1), 0);
        }
    }
    (0, 0)
}

/// The background analysis loop.
fn analysis_loop(
    rx: Receiver<AnalysisRequest>,
    tx: Sender<AnalysisEvent>,
    latest_version: Arc<AtomicU64>,
) {
    let mut pending: Option<(String, String, u64, Instant)> = None;

    loop {
        let timeout = pending.as_ref().map(|(_, _, _, at)| {
            let elapsed = at.elapsed();
            let debounce = Duration::from_millis(DEBOUNCE_MS);
            if elapsed >= debounce {
                Duration::ZERO
            } else {
                debounce - elapsed
            }
        });

        let recv_result = if let Some(t) = timeout {
            if t.is_zero() {
                rx.try_recv()
                    .map_err(|_| crossbeam_channel::RecvTimeoutError::Timeout)
            } else {
                rx.recv_timeout(t)
            }
        } else {
            rx.recv()
                .map_err(|_| crossbeam_channel::RecvTimeoutError::Disconnected)
        };

        match recv_result {
            Ok(AnalysisRequest::Shutdown) => return,
            Ok(AnalysisRequest::Update {
                uri,
                content,
                version,
            }) => {
                pending = Some((uri, content, version, Instant::now()));
                // Drain queued updates to coalesce
                while let Ok(req) = rx.try_recv() {
                    match req {
                        AnalysisRequest::Shutdown => return,
                        AnalysisRequest::Update {
                            uri,
                            content,
                            version,
                        } => {
                            pending = Some((uri, content, version, Instant::now()));
                        }
                    }
                }
            }
            Err(crossbeam_channel::RecvTimeoutError::Timeout) => {}
            Err(crossbeam_channel::RecvTimeoutError::Disconnected) => return,
        }

        // Check if debounce period elapsed
        if let Some((ref uri, ref content, version, at)) = pending {
            if at.elapsed() >= Duration::from_millis(DEBOUNCE_MS) {
                let current = latest_version.load(Ordering::Relaxed);
                if version >= current {
                    let lang = extract::detect_language(uri);
                    let keywords = extract::language_keywords(lang);

                    let (symbols, diagnostics) = parse_content(lang, content);

                    tx.send(AnalysisEvent::Analysis(AnalysisResult {
                        uri: uri.clone(),
                        version,
                        symbols,
                        diagnostics,
                        keywords,
                    }))
                    .ok();
                }
                pending = None;
            }
        }
    }
}
