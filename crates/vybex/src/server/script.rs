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
    timeout_secs: u64,
    shutdown: Option<Arc<tokio::sync::Notify>>,
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

    // Build the hyper response from the streaming channel. Wrap in a
    // race between (a) the configured per-request timeout and (b) the
    // server-wide shutdown notify. Whichever fires first releases the
    // handler with an error response so Ctrl+C doesn't get stuck behind
    // hung scripts.
    let start = std::time::Instant::now();
    let deadline = std::time::Duration::from_secs(timeout_secs);

    enum Outcome {
        Done(Response<BoxBody>),
        Timeout,
        Shutdown,
    }

    let outcome = {
        let shutdown_fut = async {
            if let Some(n) = shutdown.as_ref() {
                n.notified().await;
            } else {
                // No shutdown wired (programmatic / test path): park forever.
                std::future::pending::<()>().await;
            }
        };
        tokio::select! {
            resp = build_response(response_rx) => Outcome::Done(resp),
            _ = tokio::time::sleep(deadline) => Outcome::Timeout,
            _ = shutdown_fut => Outcome::Shutdown,
        }
    };

    match outcome {
        Outcome::Done(r) => r,
        Outcome::Timeout => {
            let elapsed = start.elapsed();
            eprintln!(
                "[vybex] ERROR: script timeout after {:.2}s (configured timeout_secs={}) — script did not emit any response. Script: {}  Likely cause: infinite loop, blocked await, or missing host function. The VM thread is still running; the HTTP response is going out as 504 now.",
                elapsed.as_secs_f64(),
                timeout_secs,
                script_path.display(),
            );
            let body = format!(
                "504 Gateway Timeout\n\nThe script {:?} did not emit a response within {}s.\n\nThis usually means:\n  - an infinite loop in the script\n  - a blocked host-function call (unreachable WASI await)\n  - a missing import that left a value undefined and the script is re-trying\n\nServer stderr has the script path and elapsed time. The VM worker thread is orphaned (will run until it naturally exits); restart the server if this recurs.\n",
                script_path.display(),
                timeout_secs,
            );
            bytes_response(504, "text/plain; charset=utf-8", body.into_bytes())
        }
        Outcome::Shutdown => {
            let elapsed = start.elapsed();
            eprintln!(
                "[vybex] WARN: script aborted by shutdown after {:.2}s. Script: {}  (the VM worker thread is orphaned but the process is exiting anyway)",
                elapsed.as_secs_f64(),
                script_path.display(),
            );
            let body = format!(
                "503 Service Unavailable\n\nServer is shutting down. Request to {:?} was aborted after {:.2}s.\n",
                script_path.display(),
                elapsed.as_secs_f64(),
            );
            bytes_response(503, "text/plain; charset=utf-8", body.into_bytes())
        }
    }
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
        // Directory-serving scripts need local filesystem access for
        // PHP/webroot routing (`is_dir`, `file_exists`, includes, etc.).
        // The host layer registers the current filesystem surface as a
        // single module behind FileRead/FileWrite gating, so FileRead is
        // the narrowest capability that makes those imports resolvable.
        let mut c = Capabilities::safe();
        c.grant(Capability::FileRead);
        c.grant(Capability::HttpServer);
        c
    };
    vybe_host::register_with_capabilities(&mut vm, &caps);

    // Populate PHP-style superglobals from the request context. Built as
    // `ObjectKind::Map` — the canonical cross-language associative type
    // — so any language's string-key access via `ecma:array.get` works
    // uniformly. PHP's `$_SERVER['REQUEST_METHOD']` lands here.
    inject_superglobals(&mut vm, &ctx);

    // SAPI-style output override: re-register `wasi:cli/log` (what PHP `echo`,
    // JS `console.log`, and most language `print` calls compile to) to write
    // to the HTTP response body when a request context is installed. Mirrors
    // how PHP's real `sapi_module->ub_write` is swapped per SAPI. The
    // default stderr path stays for anything that reaches the fn with no
    // context (e.g., inside a callback on a bare thread).
    vm.register_host_fn("wasi:cli", "log", Box::new(|_ctx, args| {
        // PHP echo emits one call per argument with arity 1, so we don't
        // join with spaces here — each call writes its single arg verbatim.
        // Semantics: "no newline, no joining" matches real PHP `echo`.
        let mut buf = Vec::<u8>::new();
        for a in args {
            match a {
                vybe_bytecode::Value::String(s) => buf.extend_from_slice(s.as_bytes()),
                other => buf.extend_from_slice(format!("{}", other).as_bytes()),
            }
        }
        match vybe_host::with_context(|c| {
            c.response.lock().unwrap().write_bytes(buf.clone());
        }) {
            Some(()) => {} // wrote to response
            None => {
                // CLI fallback — mirrors original console::register behavior
                // (println per log call).
                let parts: Vec<String> = args.iter().map(|v| format!("{}", v)).collect();
                println!("{}", parts.join(" "));
            }
        }
        vybe_bytecode::Value::Null
    }));

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

/// Populate PHP-style superglobals by inserting entries into `vm.globals`.
///
/// Each superglobal is an `ObjectKind::Map` (the canonical cross-language
/// associative type). `$_SERVER['REQUEST_METHOD']` then routes through
/// `ecma:array.get` which dispatches on `ObjectKind::Map` and returns
/// the value — same as every other language's associative map.
///
/// The globals inserted here are PHP-idiomatic (`_SERVER`, `_GET`, etc.)
/// but non-PHP scripts running under `--serve` simply won't touch them.
/// Real request data is always available via `\Vybe\Http\Request\*` host
/// calls regardless of language.
fn inject_superglobals(vm: &mut vybe_bytecode::VM, ctx: &Arc<RequestContext>) {
    use indexmap::IndexMap;
    use vybe_bytecode::value::{Object, ObjectKind, Value};
    use std::sync::{Arc as StdArc, Mutex as StdMutex};

    fn make_map(pairs: impl IntoIterator<Item = (String, String)>) -> Value {
        let mut im = IndexMap::new();
        for (k, v) in pairs {
            im.insert(
                Value::String(StdArc::from(k.as_str())),
                Value::String(StdArc::from(v.as_str())),
            );
        }
        let mut obj = Object::new();
        obj.kind = ObjectKind::Map(im);
        Value::Object(StdArc::new(StdMutex::new(obj)))
    }

    let server = make_map(ctx.env.iter().map(|(k, v)| (k.clone(), v.clone())));

    let get_pairs: Vec<(String, String)> = form_urlencoded::parse(ctx.query.as_bytes())
        .map(|(k, v)| (k.into_owned(), v.into_owned()))
        .collect();
    let get = make_map(get_pairs);

    let cookie_header = ctx.headers.iter()
        .find(|(n, _)| n.eq_ignore_ascii_case("cookie"))
        .map(|(_, v)| v.as_str())
        .unwrap_or("");
    let cookie_pairs = parse_cookie_header(cookie_header);
    let cookies = make_map(cookie_pairs);

    // $_POST — populated only for form-urlencoded bodies. multipart/form-data
    // handling (with $_FILES) lands in a follow-up.
    let post = {
        let content_type = ctx.headers.iter()
            .find(|(n, _)| n.eq_ignore_ascii_case("content-type"))
            .map(|(_, v)| v.as_str())
            .unwrap_or("");
        if content_type.to_ascii_lowercase().starts_with("application/x-www-form-urlencoded") {
            let body = ctx.body.lock().unwrap().read_all();
            let pairs: Vec<(String, String)> = form_urlencoded::parse(&body)
                .map(|(k, v)| (k.into_owned(), v.into_owned()))
                .collect();
            make_map(pairs)
        } else {
            make_map(Vec::<(String, String)>::new())
        }
    };

    // PHP variables and functions live in separate namespaces; the
    // walker preserves the `$` sigil on variable identifiers so a
    // function `foo` and a variable `$foo` don't collide. Register
    // the superglobals with the same `$` prefix so user code's
    // `$_SERVER["PHP_SELF"]` etc. resolves correctly.
    vm.globals.insert("$_SERVER".to_string(), server);
    vm.globals.insert("$_GET".to_string(), get);
    vm.globals.insert("$_COOKIE".to_string(), cookies);
    vm.globals.insert("$_POST".to_string(), post);
    let mut request_im: IndexMap<Value, Value> = IndexMap::new();
    for key in ["$_GET", "$_POST", "$_COOKIE"] {
        if let Some(Value::Object(obj)) = vm.globals.get(key) {
            let o = obj.lock().unwrap();
            if let ObjectKind::Map(ref im) = o.kind {
                for (k, v) in im.iter() {
                    request_im.insert(k.clone(), v.clone());
                }
            }
        }
    }
    let mut req_obj = Object::new();
    req_obj.kind = ObjectKind::Map(request_im);
    vm.globals.insert(
        "$_REQUEST".to_string(),
        Value::Object(StdArc::new(StdMutex::new(req_obj))),
    );
}

fn parse_cookie_header(header: &str) -> Vec<(String, String)> {
    let mut out = Vec::new();
    for part in header.split(';') {
        let part = part.trim();
        if part.is_empty() { continue; }
        match part.split_once('=') {
            Some((n, v)) => {
                let v = v.strip_prefix('"').and_then(|s| s.strip_suffix('"')).unwrap_or(v);
                out.push((n.trim().to_string(), v.to_string()));
            }
            None => out.push((part.to_string(), String::new())),
        }
    }
    out
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
