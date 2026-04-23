//! Compile + run a script per request.
//!
//! Phase 1: fresh VM per request. Compile cache + VM pool are Phase 2.
//! The script runs on a `spawn_blocking` worker with the thread-local
//! `REQUEST_CONTEXT` installed; all `vybe:http/*` host fns read from it.
//! Whatever the script writes through `vybe:http/response.*` flows out
//! through the `ResponseMessage` channel owned by the `RequestContext`
//! and is assembled into a hyper response by `response_stream`.
//!
//! Languages that don't yet have a server adapter (everything besides
//! PHP in Phase 1) will get their CLI-style output. The PHP adapter
//! that routes `echo` through `vybe:http/response.write` lands in the
//! next slice.

use std::path::{Path, PathBuf};
use std::sync::Arc;

use super::response_stream::{bytes_response, build_response, BoxBody};
use bytes::Bytes;
use http::Response;
use vybe_host::{Capabilities, Capability, RequestContext};

pub async fn serve(
    script_path: PathBuf,
    ctx: Arc<RequestContext>,
    response_rx: std::sync::mpsc::Receiver<vybe_host::ResponseMessage>,
    no_sandbox: bool,
) -> Response<BoxBody> {
    // Kick off the VM on a blocking worker. We don't await it here —
    // the response stream bridge will await messages on response_rx as
    // they arrive. When the VM is done, the sender is dropped and the
    // body stream closes naturally.
    let script = script_path.clone();
    let vm_ctx = Arc::clone(&ctx);
    tokio::task::spawn_blocking(move || {
        run_vm(&script, vm_ctx, no_sandbox);
    });

    // Build the hyper response from the streaming channel. This awaits
    // the first message (Headers) before returning, so we have proper
    // status + headers before any bytes hit the wire.
    build_response(response_rx).await
}

fn run_vm(script_path: &Path, ctx: Arc<RequestContext>, no_sandbox: bool) {
    use vybe_bytecode::VM;

    // Install the thread-local context for the duration of this VM run.
    let _guard = vybe_host::install_context(Arc::clone(&ctx));

    // Compile the script (Phase 2: compile cache).
    let bundle = match crate::projects::load(script_path) {
        Ok(b) => b,
        Err(e) => {
            let msg = format!("compile error: {e}");
            end_with_text(&ctx, 500, &msg);
            return;
        }
    };
    let chunks = match bundle.compile() {
        Ok(c) => c,
        Err(e) => {
            let msg = format!("compile error: {e}");
            end_with_text(&ctx, 500, &msg);
            return;
        }
    };

    // Fresh VM per request (Phase 2: pool).
    let mut vm = VM::new();
    let caps = if no_sandbox {
        Capabilities::all()
    } else {
        // Safe + HttpServer so the script can read req / write resp.
        let mut c = Capabilities::safe();
        c.grant(Capability::HttpServer);
        c
    };
    vybe_host::register_with_capabilities(&mut vm, &caps);

    // Register WASM-named globals (mirrors main.rs). Allows scripts to
    // reference named functions across multi-file projects.
    for (idx, chunk) in chunks.iter().enumerate() {
        if !chunk.name.is_empty()
            && chunk.name != "<script>"
            && chunk.name != "<bootstrap>"
            && !chunk.name.starts_with("__stdlib_")
        {
            let func = vybe_bytecode::value::Function {
                name: Some(chunk.name.clone()),
                arity: chunk.arity,
                chunk_index: idx,
                upvalues: vec![],
            };
            let mut obj = vybe_bytecode::value::Object::new();
            obj.kind = vybe_bytecode::value::ObjectKind::Function(func);
            let val = vybe_bytecode::Value::Object(std::sync::Arc::new(std::sync::Mutex::new(obj)));
            vm.globals.insert(chunk.name.to_lowercase(), val);
        }
    }

    if let Err(e) = vm.run(chunks) {
        // If the response hasn't been flushed yet, we can still return
        // a proper 500. Otherwise we can only log; headers are gone.
        let headers_sent = ctx.response.lock().unwrap().headers_sent;
        if !headers_sent {
            let msg = format!("runtime error: {e}");
            end_with_text(&ctx, 500, &msg);
            return;
        } else {
            eprintln!("[vybex] runtime error after response started: {e}");
        }
    }

    // Ensure end() is called so the client sees EOF, even if the script
    // forgot.
    ctx.response.lock().unwrap().end();
}

fn end_with_text(ctx: &RequestContext, status: u16, body: &str) {
    let mut r = ctx.response.lock().unwrap();
    if !r.headers_sent {
        r.status = status;
        r.headers.retain(|(n, _)| !n.eq_ignore_ascii_case("content-type"));
        r.headers.push((
            "Content-Type".to_string(),
            "text/plain; charset=utf-8".to_string(),
        ));
    }
    r.write_bytes(body.as_bytes().to_vec());
    r.end();
}

/// Fallback response builder for when we want to short-circuit without
/// running a script at all (e.g., unsupported extension, stub for
/// languages without an adapter yet).
#[allow(dead_code)]
pub fn not_implemented(body: &str) -> Response<BoxBody> {
    bytes_response(501, "text/plain; charset=utf-8", body.as_bytes().to_vec())
}

// Silence unused import on Bytes in the signature-only path.
#[allow(dead_code)]
fn _bytes_shim() -> Bytes { Bytes::new() }
