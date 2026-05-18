//! Programmatic mode — `server.listen(addr, handler)` primitive for
//! Node/Flask/Sinatra-style scripts.
//!
//! Model:
//!
//!   Script thread                         Tokio thread
//!   ─────────────                         ────────────
//!   vm.run(chunks)
//!     ↓
//!   server.listen(addr, handler)   ─→     std::thread::spawn
//!     │                                    │
//!     │                                    tokio::runtime + hyper
//!     │                                    per-request:
//!     │                                    ┌─ build RequestContext
//!     │  ◄── mpsc::Sender<Event> ──        ├─ send Event(ctx, resp_tx)
//!     pulls Event from channel             │
//!     │                                    │
//!     install_context(ctx)                 │
//!     ctx.invoke(handler)                  │
//!       └ script handler runs,             │
//!         reads vybe:http/request.*,       │
//!         writes vybe:http/response.*      │
//!     response channel drains     ─→       │  (response streams out)
//!     oneshot .send(()) done        ─→    resp_tx fires
//!                                          writes final frames, closes conn
//!
//! A single VM thread processes requests serially (Node.js event-loop
//! model). Concurrency is async-I/O on the tokio side; CPU-bound work
//! in handlers blocks the next request. Future work: JSPI integration
//! for handler-level await semantics.
//!
//! Uses real WASI/hyper primitives only — no custom vybe:* shortcuts.

use std::net::SocketAddr;
use std::sync::Arc;
use std::sync::mpsc as std_mpsc;

use bytes::Bytes;
use http::Request;
use http_body_util::BodyExt;
use hyper::body::Incoming;
use vybe_bytecode::{VM, Value, HostContext};
use vybe_host::{install_context, RequestContext, ResponseMessage};

use super::response_stream::{BoxBody, build_response};

/// Register the `vybe:http/server.listen` host fn on `vm`. Call once
/// at VM setup time (e.g. before `vm.run`) so the primitive is
/// available to any script that wants to act as a server.
///
/// Programmatic mode deliberately does NOT override `wasi:cli/log`.
/// That would be a Layer-2 shortcut bypassing Layer 3 (per-language
/// adapter stdlib). Node cleanly separates `console.log` (→ stdout,
/// server log) from `res.write` (→ HTTP response body). The JS stdlib
/// `node:http` (Phase 2 deliverable) provides `ServerResponse.write`
/// which forwards to `vybe:http/response.write`. Scripts that want to
/// write to the response today call `Vybe.Http.Response.write(...)`
/// directly via Component Model qualified names — until the Layer 3
/// JS stdlib ships.
///
/// Directory mode (`--serve`) is different: PHP's SAPI semantics
/// legitimately route `echo` to the response body, and that override
/// lives in `script.rs` where it applies to the PHP-driven model.
pub fn register(vm: &mut VM) {
    vm.register_host_fn("node:http", "listen", Box::new(listen_host_fn));
    vm.register_host_fn("node:http", "close", Box::new(|_ctx, _args| Value::Null));
    vm.register_host_fn("node:http", "address", Box::new(|_ctx, _args| Value::Null));
}

/// Event pushed from the tokio thread to the VM thread per request.
struct ServerEvent {
    ctx: Arc<RequestContext>,
    // Signal back to the tokio side that the handler returned. Response
    // bytes stream separately via the RequestContext's response channel.
    done_tx: tokio::sync::oneshot::Sender<()>,
}

fn listen_host_fn(ctx: &mut HostContext, args: &[Value]) -> Value {
    let addr_raw = match args.first() {
        Some(Value::String(s)) => s.to_string(),
        Some(other) => format!("{}", other),
        None => return Value::Null,
    };
    let addr_normalized: String = addr_raw
        .strip_prefix(':')
        .map(|p| format!("127.0.0.1:{p}"))
        .unwrap_or(addr_raw);
    let bind_addr: SocketAddr = match addr_normalized.parse() {
        Ok(a) => a,
        Err(e) => {
            eprintln!("[vybex] server.listen: invalid address {addr_normalized:?}: {e}");
            return Value::Null;
        }
    };

    let handler = args.get(1).cloned().unwrap_or(Value::Null);

    let (event_tx, event_rx) = std_mpsc::channel::<ServerEvent>();

    // Spawn the hyper + tokio runtime on a background thread. It pushes
    // ServerEvents into event_tx for the VM thread to consume serially.
    let tokio_thread = std::thread::Builder::new()
        .name("vybex-programmatic-server".into())
        .spawn(move || {
            let rt = match tokio::runtime::Builder::new_multi_thread()
                .enable_all()
                .build()
            {
                Ok(rt) => rt,
                Err(e) => {
                    eprintln!("[vybex] server.listen: failed to build tokio runtime: {e}");
                    return;
                }
            };
            rt.block_on(async move {
                let listener = match tokio::net::TcpListener::bind(bind_addr).await {
                    Ok(l) => l,
                    Err(e) => {
                        eprintln!("[vybex] server.listen: bind {bind_addr} failed: {e}");
                        return;
                    }
                };
                eprintln!("[vybex] listening on http://{bind_addr}");

                let local_addr = listener.local_addr().unwrap_or(bind_addr);
                let event_tx = event_tx;

                loop {
                    let (stream, remote) = match listener.accept().await {
                        Ok(p) => p,
                        Err(e) => { eprintln!("[vybex] accept error: {e}"); continue; }
                    };
                    let event_tx = event_tx.clone();
                    tokio::spawn(async move {
                        let io = hyper_util::rt::TokioIo::new(stream);
                        let svc = hyper::service::service_fn(move |req: Request<Incoming>| {
                            let event_tx = event_tx.clone();
                            async move {
                                Ok::<_, std::convert::Infallible>(
                                    handle_request(req, remote, local_addr, event_tx).await
                                )
                            }
                        });
                        let builder = hyper_util::server::conn::auto::Builder::new(
                            hyper_util::rt::TokioExecutor::new(),
                        );
                        if let Err(e) = builder.serve_connection(io, svc).await {
                            eprintln!("[vybex] connection error: {e}");
                        }
                    });
                }
            });
        });

    if let Err(e) = tokio_thread {
        eprintln!("[vybex] server.listen: failed to spawn runtime thread: {e}");
        return Value::Null;
    }

    // Main thread event loop — pull events, invoke handler per request.
    // Blocks forever (until Ctrl+C or channel close). Since the tokio
    // thread holds the only other sender and runs forever, channel
    // close would mean the runtime died.
    while let Ok(event) = event_rx.recv() {
        let ServerEvent { ctx: req_ctx, done_tx } = event;
        {
            let _guard = install_context(Arc::clone(&req_ctx));
            ctx.invoke(&handler, &[]);
            // If the handler didn't call response.end(), do it implicitly
            // so the body stream can close.
            req_ctx.response.lock().unwrap().end();
        }
        // Ignore send-failure — means the client hung up; fine.
        let _ = done_tx.send(());
    }

    Value::Null
}

async fn handle_request(
    req: Request<Incoming>,
    remote: SocketAddr,
    local_addr: SocketAddr,
    event_tx: std_mpsc::Sender<ServerEvent>,
) -> http::Response<BoxBody> {
    // Collect body up to 10 MiB (same default as --serve; tune later).
    const MAX_BODY: usize = 10 * 1024 * 1024;
    let (parts, body) = req.into_parts();
    let body_bytes = match read_limited(body, MAX_BODY).await {
        Ok(b) => b,
        Err(e) => {
            return super::response_stream::bytes_response(
                413,
                "text/plain; charset=utf-8",
                format!("body error: {e}\n").into_bytes(),
            );
        }
    };

    let reassembled = http::Request::from_parts(parts, ());
    let built = super::request_context::build(
        &reassembled,
        body_bytes,
        remote,
        None,           // no script filename — user's handler *is* the entry
        None,
        std::path::Path::new("."),
        local_addr,
        "http",
    );

    // Hand the request to the VM thread, await handler completion. The
    // body channel inside RequestContext streams out independently —
    // build_response reads it and produces the hyper body.
    let (done_tx, done_rx) = tokio::sync::oneshot::channel::<()>();
    let event = ServerEvent { ctx: Arc::clone(&built.ctx), done_tx };
    if event_tx.send(event).is_err() {
        return super::response_stream::bytes_response(
            503,
            "text/plain; charset=utf-8",
            b"server thread unavailable\n".to_vec(),
        );
    }

    // Build the hyper response from the RequestContext's response
    // channel. This returns as soon as the first Headers message
    // arrives, and streams subsequent Data frames. It does NOT wait
    // for done_rx — body flows while the handler runs.
    let resp = build_response(built.response_rx).await;

    // Spawn a task to await done_tx so we don't leak it; we don't
    // actually need to block on it here — the response body drives
    // completion via the EOF of the streaming channel.
    tokio::spawn(async move { let _ = done_rx.await; });

    resp
}

async fn read_limited(mut body: Incoming, max: usize) -> Result<Bytes, String> {
    let mut buf = Vec::<u8>::new();
    while let Some(frame) = body.frame().await {
        let frame = frame.map_err(|e| e.to_string())?;
        if let Some(data) = frame.data_ref() {
            if buf.len() + data.len() > max {
                return Err(format!("body exceeds {max} bytes"));
            }
            buf.extend_from_slice(data);
        }
    }
    Ok(Bytes::from(buf))
}

// Suppress unused-import warning on ResponseMessage when the compiler
// strips inline uses; it's load-bearing via the types in request_context
// and response_stream signatures.
#[allow(dead_code)]
fn _type_witness() -> Option<ResponseMessage> { None }
