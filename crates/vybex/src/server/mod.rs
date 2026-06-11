//! `vybex --serve` — HTTP server for Vybe.
//!
//! See `httpserver.md` at the repo root for the full architecture. This
//! module is the Layer 2 glue: a `hyper`-based HTTP/1.1 server that
//! dispatches requests to either the static-file fast path or a script
//! execution via `spawn_blocking`. Language-specific web idioms
//! (superglobals, Node `(req, res)`, WSGI, Rack, ASP.NET HttpContext)
//! live in each language's own stdlib on top of the `vybe:http` host
//! module — this file does NOT know PHP from Python.

pub mod config;
pub mod directory;
pub mod errors;
pub mod hyper_service;
pub mod logging;
pub mod programmatic;
pub mod request_context;
pub mod response_stream;
pub mod script;
pub mod static_files;

pub use config::ServeConfig;

use std::net::SocketAddr;

/// Entry point for directory mode. Called from `main.rs` when `--serve`
/// is passed. Never returns — the tokio runtime parks until SIGINT.
pub fn serve_directory(config: ServeConfig) -> ! {
    // 32 MB blocking-thread stack — PHP compilation of large combined files
    // (many require_once classes) uses significant Rust stack depth and overflows
    // the default 2 MB tokio blocking-worker stack on complex requests.
    let rt = tokio::runtime::Builder::new_multi_thread()
        .enable_all()
        .thread_name("vybex-serve")
        .thread_stack_size(32 * 1024 * 1024)
        .build()
        .expect("failed to build tokio runtime");

    rt.block_on(async move {
        if let Err(e) = run(config).await {
            eprintln!("[vybex] server error: {e}");
            std::process::exit(1);
        }
    });

    std::process::exit(0);
}

async fn run(config: ServeConfig) -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
    // Allow `:PORT` as a shortcut for `127.0.0.1:PORT` (like `go run` / tiny-http convention).
    let bind_normalized: String = if let Some(port) = config.bind.strip_prefix(':') {
        format!("127.0.0.1:{port}")
    } else {
        config.bind.clone()
    };
    let addr: SocketAddr = bind_normalized
        .parse()
        .map_err(|e| format!("invalid bind address {:?}: {e}", config.bind))?;

    let listener = tokio::net::TcpListener::bind(addr).await?;

    eprintln!(
        "[vybex] serving {} on http://{}",
        config.root.display(),
        addr
    );
    eprintln!("[vybex] press Ctrl+C to stop");

    // Shutdown notification shared with every in-flight request handler.
    // On Ctrl+C we flip this; per-request timeouts race against it so hung
    // scripts are released with a 503 immediately instead of blocking the
    // drain for up to `timeout_secs`.
    let shutdown = std::sync::Arc::new(tokio::sync::Notify::new());
    let config_with_shutdown = {
        let mut c = config;
        c.shutdown = Some(std::sync::Arc::clone(&shutdown));
        c
    };
    let shared = std::sync::Arc::new(config_with_shutdown);

    // Graceful shutdown: drain in-flight on Ctrl+C.
    let graceful = hyper_util::server::graceful::GracefulShutdown::new();
    let mut shutdown_signal = Box::pin(tokio::signal::ctrl_c());

    loop {
        tokio::select! {
            accept = listener.accept() => {
                let (stream, remote) = match accept {
                    Ok(pair) => pair,
                    Err(e) => { eprintln!("[vybex] accept error: {e}"); continue; }
                };
                let io = hyper_util::rt::TokioIo::new(stream);
                let svc = hyper_service::make_service(std::sync::Arc::clone(&shared), remote);
                let builder = hyper_util::server::conn::auto::Builder::new(hyper_util::rt::TokioExecutor::new());
                let conn = builder.serve_connection(io, svc);
                let fut = graceful.watch(conn.into_owned());
                tokio::spawn(async move {
                    if let Err(e) = fut.await {
                        eprintln!("[vybex] connection error: {e}");
                    }
                });
            }
            _ = &mut shutdown_signal => {
                eprintln!("[vybex] Ctrl+C received, aborting in-flight requests…");
                shutdown.notify_waiters();
                break;
            }
        }
    }

    // Keep the drain tight. Any request that hasn't already released on
    // the shutdown notify gets ~2s before we hard-exit. Dev servers
    // prioritise fast Ctrl+C over graceful completion.
    tokio::select! {
        _ = graceful.shutdown() => {
            eprintln!("[vybex] all connections drained");
        }
        _ = tokio::time::sleep(std::time::Duration::from_secs(2)) => {
            eprintln!("[vybex] drain timeout after 2s; forcing exit");
        }
    }

    Ok(())
}
