//! Compile + run a script per request.
//!
//! Two execution paths, same per-request body:
//!
//! - **warm** (default) — the request is handed to a [`vm_pool`](super::vm_pool)
//!   thread that booted its VM once and `reset_to`s it between tenants. Only
//!   [`run_request`] runs per request.
//! - **cold** (`--cold`) — a fresh `VM::new()` per request, via [`run_vm`].
//!   Kept as the control to diff the warm path against, and as the escape
//!   hatch if a leak ever turns up in the field.
//!
//! Both paths use the [compile cache](super::compile_cache), and `--no-cache`
//! turns it off for both. The two flags are deliberately separate: `--cold`
//! isolates the VM pool, `--no-cache` isolates the cache, and one flag doing
//! both would leave neither measurable on its own.
//!
//! The script runs on a blocking thread with the thread-local
//! `REQUEST_CONTEXT` installed; all `vybe:http/*` host fns read from it.
//! Whatever the script writes through `vybe:http/response.*` flows out
//! through the `ResponseMessage` channel owned by the `RequestContext`
//! and is assembled into a hyper response by `response_stream`.
//!
//! CLI-style stdout from scripts is bound to the HTTP response body for
//! request handling. Language-level buffering (PHP `ob_*`, etc.) happens
//! before bytes reach the WASI stdout stream.

use std::path::{Path, PathBuf};
use std::sync::Arc;

use super::response_stream::{BoxBody, build_response, bytes_response};
use bytes::Bytes;
use http::Response;
use indexmap::IndexMap;
use vybe_platform_node::http::RequestContext;
use vybe_runtime::capabilities::{Capabilities, Capability};

const PHP_SESSION_COOKIE_NAME: &str = "PHPSESSID";
const PHP_SESSION_ID_GLOBAL: &str = "__php_session_id";
const PHP_SESSION_STARTED_GLOBAL: &str = "__php_session_started";
const PHP_SESSION_NEEDS_COOKIE_GLOBAL: &str = "__php_session_needs_cookie";
const PHP_SESSION_DESTROYED_GLOBAL: &str = "__php_session_destroyed";

static PHP_SESSION_STORE: std::sync::LazyLock<
    dashmap::DashMap<String, IndexMap<String, vybe_runtime::Value>>,
> = std::sync::LazyLock::new(dashmap::DashMap::new);

#[allow(clippy::too_many_arguments)]
pub async fn serve(
    script_path: PathBuf,
    ctx: Arc<RequestContext>,
    response_rx: std::sync::mpsc::Receiver<vybe_platform_node::http::ResponseMessage>,
    no_sandbox: bool,
    timeout_secs: u64,
    shutdown: Option<Arc<tokio::sync::Notify>>,
    pool: Option<Arc<super::vm_pool::VmPool>>,
    cache: Option<Arc<super::compile_cache::CompileCache>>,
) -> Response<BoxBody> {
    // Kick off the VM. We don't await it here — the response stream bridge
    // will await messages on response_rx as they arrive. When the VM is done,
    // the sender is dropped and the body stream closes naturally.
    let script = script_path.clone();
    let vm_ctx = Arc::clone(&ctx);
    match pool {
        Some(pool) => {
            let job = super::vm_pool::Job {
                script,
                ctx: vm_ctx,
            };
            if let Err(e) = pool.submit(job) {
                return bytes_response(
                    500,
                    "text/plain; charset=utf-8",
                    format!("500 Internal Server Error\n\n{e}\n").into_bytes(),
                );
            }
        }
        // `--cold`: the pre-pool path, one whole VM per request.
        None => {
            tokio::task::spawn_blocking(move || {
                run_vm(&script, vm_ctx, no_sandbox, cache.as_ref());
            });
        }
    }

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

/// The capability set a served script runs with.
///
/// Lives here rather than inline in `run_vm` because the warm pool has to
/// build it ONCE at boot — a pooled VM's capabilities are fixed for its life —
/// and a second copy of this reasoning would be a sandbox that differs between
/// `--cold` and the default path.
pub fn caps_for(no_sandbox: bool) -> Capabilities {
    if no_sandbox {
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
        // A served script READS the request it was handed and WRITES its
        // response — both are `wasi:http`, which the wasi plugin gates on
        // `Capability::Http`. Without this every superglobal dies with
        // `Unresolved import: "wasi:http/types" …`, because the request is no
        // longer mirrored into Rust-side globals. Note this grants no outbound
        // reach on its own: making a request needs `outgoing-handler`, and the
        // sandbox still withholds `Sockets`.
        c.grant(Capability::Http);
        // `$_ENV` / `getenv()` are `wasi:cli/environment`.
        c.grant(Capability::Environment);
        c
    }
}

/// SAPI-style stdout override: bind the WASI stdout stream to the HTTP
/// response body. PHP `echo`/`print` and other normal script output should use
/// stdout; diagnostics stay on logging/stderr.
///
/// Registered ONCE per warm VM, not once per request. It already reads the
/// `RequestContext` out of the thread-local rather than closing over one — the
/// closure was written request-agnostic from the start — so the only thing
/// that ever made it per-request was living inside `run_vm`.
pub fn register_response_stdout(vm: &mut vybe_runtime::VM) {
    vm.register_host_fn(
        "wasi:cli/stdout",
        "write-via-stream",
        Box::new(|host_ctx, args| {
            let stream_val = args.first().cloned().unwrap_or(vybe_runtime::Value::Null);
            let bytes = host_ctx.stream_drain(&stream_val);
            if !bytes.is_empty() {
                match vybe_platform_node::http::with_context(|c| {
                    c.response.lock().unwrap().write_bytes(bytes.clone());
                }) {
                    Some(()) => {}
                    None => {
                        use std::io::Write;
                        let _ = std::io::stdout().write_all(&bytes);
                    }
                }
            }
            let (fut, fut_id) = host_ctx.create_future();
            host_ctx.resolve_future(fut_id, vybe_runtime::Value::Null);
            fut
        }),
    );
}

/// `--cold`: one whole VM for one request.
///
/// Kept as the control the warm path is diffed against. Its boot is
/// deliberately NOT `warm::boot` — no tracking, no snapshot — because a VM
/// that serves exactly one tenant has nothing to roll back to.
fn run_vm(
    script_path: &Path,
    ctx: Arc<RequestContext>,
    no_sandbox: bool,
    cache: Option<&Arc<super::compile_cache::CompileCache>>,
) {
    use vybe_runtime::VM;

    // Install the thread-local context for the duration of this VM run.
    let _guard = vybe_platform_node::http::install_context(Arc::clone(&ctx));

    let mut vm = VM::new();
    let caps = caps_for(no_sandbox);

    // Register BEFORE compiling. Language plugins publish their `LanguageDef`
    // — including the file extensions they claim — into the global registry as
    // part of this one loop, so `projects::load` can only resolve `.php` once
    // it has run. Compiling first left the registry holding nothing but the
    // built-in project types, and every served script died with
    // "Unknown file extension".
    crate::cli::register_plugins(&mut vm, &caps);
    crate::server::programmatic::register(&mut vm);
    // Ordering note: `register_all` used to run AFTER the request was
    // installed. Hoisting it above is safe only because it is documented as
    // intentionally empty today (`cli.rs`: "we do not install placeholder
    // adapters over non-existent `wasi:http/*` server surfaces"). An adapter
    // that wanted to see the request would have to move back down — and could
    // not then live in the warm baseline at all.
    if let Err(e) = crate::adapters::register_all(&mut vm) {
        let msg = format!("adapter registration error: {e}");
        end_with_text(&ctx, 500, &msg);
        return;
    }
    register_response_stdout(&mut vm);

    // The cache DOES apply here: `--cold` isolates the VM pool, `--no-cache`
    // isolates the cache. A `--cold` that also silently disabled the cache
    // would leave no way to measure either one on its own.
    run_request(&mut vm, script_path, &ctx, &caps, cache);
}

/// Everything one request costs against an already-booted VM.
///
/// This is the whole warm path: the pool thread resets its VM, installs the
/// context, and calls this. Nothing here may register a plugin, an adapter or
/// a language — those belong to the baseline, and doing them per request is
/// precisely the ~0.2s the pool exists to stop paying.
pub fn run_request(
    vm: &mut vybe_runtime::VM,
    script_path: &Path,
    ctx: &Arc<RequestContext>,
    _caps: &Capabilities,
    cache: Option<&Arc<super::compile_cache::CompileCache>>,
) {
    // `projects::load` only READS the file into a `Bundle` — no parsing happens
    // here, so there is nothing to cache at this step.
    let bundle = match vybe_compiler::projects::load(script_path) {
        Ok(b) => b,
        Err(e) => {
            let msg = format!("compile error: {e}");
            end_with_text(ctx, 500, &msg);
            return;
        }
    };

    // `wasi:http` is now the ONLY representation of the incoming request.
    // Every superglobal a PHP script reads is emitted code reading these
    // handles (`documentation/httpserver.md` §4a) — nothing below Layer 3
    // knows about PHP, WSGI or Rack. `RequestContext` survives for the
    // RESPONSE half and for the deployment metadata `wasi:http` does not model
    // (document root, resolved script path, peer address); draining those is
    // what retires it.
    install_wasi_http_request(vm, ctx);

    // Seed the session store and `$_ENV`. The request-derived superglobals are
    // NOT built here any more — see `inject_superglobals`.
    inject_superglobals(vm, ctx);

    let mut runtime_compiler = crate::dynamic::RuntimeCompilerService::new(vm);

    // The compile half below is answered from cache; the RUN half still
    // compiles every runtime `include` from source unless the service can
    // reach the same cache. Under `--serve` that is the whole per-request
    // saving given back on any page built from includes.
    if let Some(cache) = cache {
        runtime_compiler.set_include_cache(
            Arc::clone(cache) as Arc<dyn vybe_compiler::dynamic::IncludeCompileCache>
        );
    }

    // `compile_and_run_bundle` is exactly `compile_bundle` + `run_compiled`;
    // splitting it is what lets the compile half be answered from cache while
    // the run half still happens per request, against this request's globals.
    let compiled = match cache.and_then(|c| c.get(script_path)) {
        Some(Ok(hit)) => hit,
        // A cached FAILURE. Same 500 the fresh compile produced, without paying
        // for the compile again — see `Outcome::Failed`.
        Some(Err(message)) => {
            end_with_text(ctx, 500, &message);
            return;
        }
        None => {
            let (result, deps) =
                super::compile_cache::compile_with_dependencies(&mut runtime_compiler, &bundle);
            match result {
                Ok(fresh) => {
                    if let Some(c) = cache {
                        c.store(script_path, deps, Ok(&fresh));
                    }
                    fresh
                }
                Err(e) => {
                    let msg = format!("compile error: {e}");
                    if let Some(c) = cache {
                        c.store(script_path, deps, Err(msg.as_str()));
                    }
                    end_with_text(ctx, 500, &msg);
                    return;
                }
            }
        }
    };

    if let Err(e) = runtime_compiler.run_compiled(compiled) {
        // If the response hasn't been flushed yet, we can still return
        // a proper 500. Otherwise we can only log; headers are gone.
        let headers_sent = ctx.response.lock().unwrap().headers_sent;
        if !headers_sent {
            let msg = format!("runtime error: {e}");
            end_with_text(ctx, 500, &msg);
            return;
        } else {
            eprintln!("[vybex] runtime error after response started: {e}");
        }
    }

    persist_superglobals(vm, ctx);

    // Ensure end() is called so the client sees EOF, even if the script
    // forgot.
    ctx.response.lock().unwrap().end();
}

/// Seed the session into `vm.globals`.
///
/// **No superglobal is built here any more.** `$_SERVER`, `$_GET`, `$_POST`,
/// `$_FILES`, `$_COOKIE` and `$_REQUEST` are bound by the PHP walker to the
/// shared request primitives, which read `wasi:http`; `$_ENV` is bound to
/// `wasi:cli/environment.get-environment` the same way — see
/// `SUPERGLOBALS_PRELUDE` in `languages/php/src/walker.rs` and
/// `documentation/httpserver.md` §4a. This function used to re-parse the query
/// string, the `Cookie:` header and multipart bodies in Rust and read
/// `std::env::vars()` itself, which meant one request had two representations
/// and only PHP under `--serve` could ever reach the second one.
///
/// What is left is the part with no primitive behind it yet: session state
/// persists in the process-global `PHP_SESSION_STORE`, and the store needs the
/// id BEFORE the script runs in order to preload `$_SESSION`. Moving that to a
/// real backing store needs two calls that are not this change's to make (the
/// sandboxed capability set at line 133 grants only `FileRead`, and emitted
/// sessions have no end-of-request flush hook).
fn inject_superglobals(vm: &mut vybe_runtime::VM, ctx: &Arc<RequestContext>) {
    use std::sync::Arc as StdArc;
    use vybe_runtime::Value;

    let cookie_header = ctx
        .headers
        .iter()
        .find(|(n, _)| n.eq_ignore_ascii_case("cookie"))
        .map(|(_, v)| v.as_str())
        .unwrap_or("");
    let session_cookie = session_cookie_value(cookie_header);
    let session_id = session_cookie
        .clone()
        .unwrap_or_else(|| uuid::Uuid::new_v4().to_string());
    let session = make_map_value(
        PHP_SESSION_STORE
            .get(&session_id)
            .map(|entry| {
                entry
                    .iter()
                    .map(|(k, v)| (k.clone(), v.clone()))
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default(),
    );

    // PHP variables and functions live in separate namespaces; the
    // walker preserves the `$` sigil on variable identifiers so a
    // function `foo` and a variable `$foo` don't collide.
    vm.globals.insert("$_SESSION".to_string(), session);
    vm.globals.insert(
        PHP_SESSION_ID_GLOBAL.to_string(),
        Value::String(StdArc::from(session_id.as_str())),
    );
    // Publish the same identity under the NEUTRAL globals the shared session
    // primitive reads (`primitives/http_session.rs`), so `session_id()` and
    // `session_name()` — and every other language's equivalent — report the id
    // this request is actually using rather than minting a second one.
    vm.globals.insert(
        vybe_compiler::primitives::http_session::SESSION_ID_GLOBAL.to_string(),
        Value::String(StdArc::from(session_id.as_str())),
    );
    vm.globals.insert(
        vybe_compiler::primitives::http_session::SESSION_NAME_GLOBAL.to_string(),
        Value::String(StdArc::from(PHP_SESSION_COOKIE_NAME)),
    );
    vm.globals
        .insert(PHP_SESSION_STARTED_GLOBAL.to_string(), Value::Bool(false));
    vm.globals.insert(
        PHP_SESSION_NEEDS_COOKIE_GLOBAL.to_string(),
        Value::Bool(session_cookie.is_none()),
    );
    vm.globals
        .insert(PHP_SESSION_DESTROYED_GLOBAL.to_string(), Value::Bool(false));
}

/// The session cookie's value, if the request carries one.
///
/// Deliberately narrow: this is NOT a `Cookie:` header parser. `$_COOKIE` comes
/// from `common:http_cookie.request_cookies` now, and so does the session id
/// the primitive itself uses. This only answers "which stored session should
/// `$_SESSION` be preloaded from", which has to be decided in Rust because the
/// store is in Rust.
fn session_cookie_value(header: &str) -> Option<String> {
    header.split(';').find_map(|part| {
        let (name, value) = part.trim().split_once('=')?;
        if name.trim() != PHP_SESSION_COOKIE_NAME || value.is_empty() {
            return None;
        }
        Some(value.to_string())
    })
}

fn make_map_value(
    pairs: impl IntoIterator<Item = (String, vybe_runtime::Value)>,
) -> vybe_runtime::Value {
    use indexmap::IndexMap;
    use std::sync::{Arc as StdArc, Mutex as StdMutex};
    use vybe_runtime::Value;
    use vybe_runtime::value::{Object, ObjectKind};

    let mut im = IndexMap::new();
    for (k, v) in pairs {
        im.insert(Value::String(StdArc::from(k.as_str())), v);
    }
    let mut obj = Object::new();
    obj.kind = ObjectKind::Map(im);
    Value::Object(StdArc::new(StdMutex::new(obj)))
}

fn persist_superglobals(vm: &vybe_runtime::VM, _ctx: &Arc<RequestContext>) {
    let destroyed = matches!(
        vm.globals.get(PHP_SESSION_DESTROYED_GLOBAL),
        Some(vybe_runtime::Value::Bool(true))
    );

    if destroyed {
        let session_id = match vm.globals.get(PHP_SESSION_ID_GLOBAL) {
            Some(vybe_runtime::Value::String(s)) if !s.is_empty() => s.to_string(),
            _ => return,
        };
        PHP_SESSION_STORE.remove(&session_id);
        return;
    }

    let started = matches!(
        vm.globals.get(PHP_SESSION_STARTED_GLOBAL),
        Some(vybe_runtime::Value::Bool(true))
    );
    if !started {
        return;
    }

    let session_id = match vm.globals.get(PHP_SESSION_ID_GLOBAL) {
        Some(vybe_runtime::Value::String(s)) if !s.is_empty() => s.to_string(),
        _ => return,
    };

    let Some(vybe_runtime::Value::Object(obj)) = vm.globals.get("$_SESSION") else {
        return;
    };
    let guard = obj.lock().unwrap();
    let vybe_runtime::value::ObjectKind::Map(map) = &guard.kind else {
        return;
    };

    let persisted: IndexMap<String, vybe_runtime::Value> = map
        .iter()
        .map(|(key, value)| (format!("{}", key), value.clone()))
        .collect();
    PHP_SESSION_STORE.insert(session_id, persisted);
}

fn end_with_text(ctx: &RequestContext, status: u16, body: &str) {
    let mut r = ctx.response.lock().unwrap();
    if !r.headers_sent {
        r.status = status;
        r.headers
            .retain(|(n, _)| !n.eq_ignore_ascii_case("content-type"));
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
fn _bytes_shim() -> Bytes {
    Bytes::new()
}

#[cfg(test)]
mod tests {
    use super::{
        PHP_SESSION_COOKIE_NAME, PHP_SESSION_ID_GLOBAL, PHP_SESSION_STARTED_GLOBAL,
        PHP_SESSION_STORE, inject_superglobals, install_wasi_http_request, persist_superglobals,
    };
    use bytes::Bytes;
    use http::Request;
    use indexmap::IndexMap;
    use std::net::{Ipv4Addr, SocketAddr, SocketAddrV4};
    use std::path::Path;
    use std::sync::{Arc, Mutex};
    use vybe_runtime::value::ObjectKind;
    use vybe_runtime::{HostContext, VM, Value};
    fn compile_php(src: &str) -> Vec<vybe_runtime::Chunk> {
        let module = vybe_language_php::parse(src).expect("parse php");
        let profile = vybe_compiler::profile::parse_profile(vybe_language_php::profile_source())
            .expect("parse php profile");
        vybe_compiler::primitives::Compiler::with_profile(profile)
            .compile(&module)
            .expect("compile php")
    }

    fn build_ctx(
        method: &str,
        uri: &str,
        headers: &[(&str, &str)],
        body: &[u8],
    ) -> Arc<vybe_platform_node::http::RequestContext> {
        let mut builder = Request::builder().method(method).uri(uri);
        for (name, value) in headers {
            builder = builder.header(*name, *value);
        }
        let req = builder.body(()).expect("build request");
        let local = SocketAddr::V4(SocketAddrV4::new(Ipv4Addr::LOCALHOST, 8080));
        let remote = SocketAddr::V4(SocketAddrV4::new(Ipv4Addr::LOCALHOST, 54321));
        let built = crate::server::request_context::build(
            &req,
            Bytes::from(body.to_vec()),
            remote,
            Some("/tmp/index.php"),
            Some("/index.php"),
            Path::new("/tmp"),
            local,
            "http",
        );
        built.ctx
    }

    fn map_entries(value: &Value) -> IndexMap<String, Value> {
        match value {
            Value::Object(obj) => {
                let guard = obj.lock().expect("lock object");
                match &guard.kind {
                    ObjectKind::Map(map) => map
                        .iter()
                        .map(|(key, value)| (format!("{}", key), value.clone()))
                        .collect(),
                    other => panic!("expected map, got {:?}", other),
                }
            }
            other => panic!("expected object map, got {}", other),
        }
    }

    /// Read one cookie's value out of a `Set-Cookie` header.
    ///
    /// Test scaffolding, not a parser: an EMPTY value is a real answer here,
    /// because that is how `session_destroy()` clears the cookie.
    fn cookie_pair_value(header: &str, name: &str) -> Option<String> {
        header.split(';').find_map(|part| {
            let (key, value) = part.trim().split_once('=')?;
            (key.trim() == name).then(|| value.to_string())
        })
    }

    fn value_as_string(value: &Value) -> String {
        match value {
            Value::String(s) => s.to_string(),
            other => format!("{}", other),
        }
    }

    fn run_php_request_vm(
        src: &str,
        ctx: Arc<vybe_platform_node::http::RequestContext>,
    ) -> (
        Vec<String>,
        Arc<vybe_platform_node::http::RequestContext>,
        VM,
    ) {
        let chunks = compile_php(src);
        let mut vm = VM::new();
        let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
        let out = Arc::clone(&output);
        crate::cli::register_plugins(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
        vm.register_host_fn(
            "web:console",
            "log",
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let line = args
                    .iter()
                    .map(value_as_string)
                    .collect::<Vec<_>>()
                    .join(" ");
                out.lock().expect("lock output").push(line);
                Value::Null
            }),
        );

        // `echo` reaches the response through the WASI stdout stream, exactly
        // as `run_vm` binds it. Without this the harness only ever saw
        // `wasi:logging` and every script's real output was dropped on the
        // floor — assertions on `echo` compared `[]` against `[]`-shaped
        // expectations only by accident.
        let stdout_out = Arc::clone(&output);
        vm.register_host_fn(
            "wasi:cli/stdout",
            "write-via-stream",
            Box::new(move |host_ctx: &mut HostContext, args: &[Value]| {
                let stream = args.first().cloned().unwrap_or(Value::Null);
                let bytes = host_ctx.stream_drain(&stream);
                if !bytes.is_empty() {
                    stdout_out
                        .lock()
                        .expect("lock output")
                        .push(String::from_utf8_lossy(&bytes).into_owned());
                }
                let (fut, fut_id) = host_ctx.create_future();
                host_ctx.resolve_future(fut_id, Value::Null);
                fut
            }),
        );

        // Same order as `run_vm`: the request goes out through `wasi:http`
        // FIRST, because the superglobals the script reads are emitted code
        // that reads it back.
        install_wasi_http_request(&mut vm, &ctx);
        inject_superglobals(&mut vm, &ctx);
        let _guard = vybe_platform_node::http::install_context(Arc::clone(&ctx));
        vm.run(chunks).expect("run php request");
        (output.lock().expect("lock output").clone(), ctx, vm)
    }

    fn run_php_request(
        src: &str,
        ctx: Arc<vybe_platform_node::http::RequestContext>,
    ) -> (Vec<String>, Arc<vybe_platform_node::http::RequestContext>) {
        let (output, ctx, vm) = run_php_request_vm(src, ctx);
        persist_superglobals(&vm, &ctx);
        (output, ctx)
    }

    /// `$_SERVER` reaches PHP through the SHARED request primitives.
    ///
    /// Asserted from inside a running script rather than by inspecting
    /// `vm.globals`, because the binding is now emitted code — the walker's
    /// `SUPERGLOBALS_PRELUDE` calling `common:http_request.environ`, which
    /// reads `wasi:http`. Reading the global directly would pass even if the
    /// PHP name were never bound.
    #[test]
    fn server_superglobal_carries_request_and_deployment_keys() {
        let ctx = build_ctx(
            "GET",
            "http://localhost:8080/index.php?foo=bar",
            &[("Host", "localhost:8080")],
            b"",
        );
        let (out, _ctx) = run_php_request(
            r#"<?php
                echo $_SERVER['REQUEST_METHOD'];
                echo $_SERVER['QUERY_STRING'];
                echo $_SERVER['SCRIPT_NAME'];
                echo $_SERVER['HTTP_HOST'];
                echo $_SERVER['SERVER_ADDR'];
            "#,
            ctx,
        );
        assert_eq!(
            out,
            vec![
                "GET".to_string(),
                "foo=bar".to_string(),
                "/index.php".to_string(),
                "localhost:8080".to_string(),
                "127.0.0.1".to_string(),
            ],
            "message keys come from wasi:http, deployment keys from the transport"
        );
    }

    /// `$_GET` is the query string parsed by `common:http_request.query_params`.
    #[test]
    fn get_superglobal_is_parsed_from_the_query_string() {
        let ctx = build_ctx(
            "GET",
            "http://localhost:8080/index.php?name=ada&lang=php",
            &[("Host", "localhost:8080")],
            b"",
        );
        let (out, _ctx) = run_php_request(r#"<?php echo $_GET['name']; echo $_GET['lang'];"#, ctx);
        assert_eq!(out, vec!["ada".to_string(), "php".to_string()]);
    }

    /// `$_COOKIE` is the `Cookie:` header parsed by `common:http_cookie.parse`.
    #[test]
    fn cookie_superglobal_is_parsed_from_the_request_header() {
        let ctx = build_ctx(
            "GET",
            "http://localhost:8080/index.php",
            &[
                ("Host", "localhost:8080"),
                ("Cookie", "theme=dark; lang=fr"),
            ],
            b"",
        );
        let (out, _ctx) = run_php_request(
            r#"<?php echo $_COOKIE['theme']; echo $_COOKIE['lang'];"#,
            ctx,
        );
        assert_eq!(out, vec!["dark".to_string(), "fr".to_string()]);
    }

    /// A directory index resolves `SCRIPT_NAME`/`PHP_SELF` to the real file.
    ///
    /// `wasi:http` knows the request target (`/genie/`); only the transport
    /// knows it resolved to `/genie/index.php`, so this is the deployment half
    /// of the CGI environment and it must survive the merge.
    #[test]
    fn a_directory_index_reports_the_resolved_script_name() {
        let req = Request::builder()
            .method("GET")
            .uri("http://localhost:8080/genie/")
            .header("Host", "localhost:8080")
            .body(())
            .expect("build request");
        let local = SocketAddr::V4(SocketAddrV4::new(Ipv4Addr::LOCALHOST, 8080));
        let remote = SocketAddr::V4(SocketAddrV4::new(Ipv4Addr::LOCALHOST, 54321));
        let built = crate::server::request_context::build(
            &req,
            Bytes::new(),
            remote,
            Some("/tmp/genie/index.php"),
            Some("/genie/index.php"),
            Path::new("/tmp"),
            local,
            "http",
        );

        let (out, _ctx) = run_php_request(
            r#"<?php echo $_SERVER['SCRIPT_NAME']; echo $_SERVER['PHP_SELF'];"#,
            built.ctx,
        );
        assert_eq!(
            out,
            vec![
                "/genie/index.php".to_string(),
                "/genie/index.php".to_string()
            ]
        );
    }

    /// `$_POST`/`$_FILES` come from `common:http_form.*` reading the
    /// `wasi:http` body — the Rust multipart parser this replaced could only
    /// ever be reached by PHP under `--serve`.
    #[test]
    fn multipart_reaches_post_and_files() {
        let boundary = "----vybex-boundary";
        let body = format!(
            "--{boundary}\r\nContent-Disposition: form-data; name=\"title\"\r\n\r\nhello\r\n--{boundary}\r\nContent-Disposition: form-data; name=\"upload\"; filename=\"note.txt\"\r\nContent-Type: text/plain\r\n\r\nabcde\r\n--{boundary}--\r\n"
        );
        let ctx = build_ctx(
            "POST",
            "http://localhost:8080/upload.php",
            &[
                ("Host", "localhost:8080"),
                (
                    "Content-Type",
                    &format!("multipart/form-data; boundary={boundary}"),
                ),
            ],
            body.as_bytes(),
        );
        let (out, _ctx) = run_php_request(
            r#"<?php
                echo $_POST['title'];
                echo $_FILES['upload']['name'];
                echo $_FILES['upload']['type'];
                echo $_FILES['upload']['size'];
            "#,
            ctx,
        );
        assert_eq!(
            out,
            vec![
                "hello".to_string(),
                "note.txt".to_string(),
                "text/plain".to_string(),
                "5".to_string(),
            ]
        );
    }

    /// `$_REQUEST` merges the three, in PHP's documented order.
    #[test]
    fn request_superglobal_merges_get_post_and_cookie() {
        let ctx = build_ctx(
            "POST",
            "http://localhost:8080/index.php?a=fromget",
            &[
                ("Host", "localhost:8080"),
                ("Content-Type", "application/x-www-form-urlencoded"),
                ("Cookie", "c=fromcookie"),
            ],
            b"b=frompost",
        );
        let (out, _ctx) = run_php_request(
            r#"<?php echo $_REQUEST['a']; echo $_REQUEST['b']; echo $_REQUEST['c'];"#,
            ctx,
        );
        assert_eq!(
            out,
            vec![
                "fromget".to_string(),
                "frompost".to_string(),
                "fromcookie".to_string()
            ]
        );
    }

    #[test]
    fn session_start_sets_cookie_and_persists_between_requests() {
        let first_ctx = build_ctx(
            "GET",
            "http://localhost:8080/index.php",
            &[("Host", "localhost:8080")],
            b"",
        );
        let (_out1, first_ctx, first_vm) = run_php_request_vm(
            r#"<?php session_start(); $_SESSION['user'] = 'alice';"#,
            Arc::clone(&first_ctx),
        );
        assert_eq!(
            first_vm.globals.get(PHP_SESSION_STARTED_GLOBAL),
            Some(&Value::Bool(true))
        );
        let first_session = map_entries(first_vm.globals.get("$_SESSION").expect("$_SESSION"));
        assert_eq!(
            first_session.get("user").map(value_as_string),
            Some("alice".to_string())
        );
        let cookie = {
            let response = first_ctx.response.lock().expect("lock response");
            response
                .headers
                .iter()
                .find(|(name, _)| name.eq_ignore_ascii_case("set-cookie"))
                .map(|(_, value)| value.clone())
                .expect("set-cookie header")
        };
        let session_id =
            cookie_pair_value(&cookie, PHP_SESSION_COOKIE_NAME).expect("session id from cookie");
        assert_eq!(
            first_vm
                .globals
                .get(PHP_SESSION_ID_GLOBAL)
                .map(value_as_string),
            Some(session_id.clone())
        );
        persist_superglobals(&first_vm, &first_ctx);
        {
            let persisted = PHP_SESSION_STORE
                .get(&session_id)
                .expect("persisted session");
            assert_eq!(
                persisted.get("user").map(value_as_string),
                Some("alice".to_string())
            );
        }

        let second_ctx = build_ctx(
            "GET",
            "http://localhost:8080/index.php",
            &[("Host", "localhost:8080"), ("Cookie", &cookie)],
            b"",
        );
        let (out2, _second_ctx) = run_php_request(
            r#"<?php session_start(); echo $_SESSION['user'];"#,
            second_ctx,
        );
        assert_eq!(out2, vec!["alice".to_string()]);
    }

    #[test]
    fn session_destroy_removes_persisted_state_and_clears_cookie() {
        let first_ctx = build_ctx(
            "GET",
            "http://localhost:8080/index.php",
            &[("Host", "localhost:8080")],
            b"",
        );
        let (_out1, first_ctx, first_vm) = run_php_request_vm(
            r#"<?php session_start(); $_SESSION['user'] = 'alice';"#,
            Arc::clone(&first_ctx),
        );
        let cookie = {
            let response = first_ctx.response.lock().expect("lock response");
            response
                .headers
                .iter()
                .find(|(name, _)| name.eq_ignore_ascii_case("set-cookie"))
                .map(|(_, value)| value.clone())
                .expect("set-cookie header")
        };
        let session_id =
            cookie_pair_value(&cookie, PHP_SESSION_COOKIE_NAME).expect("session id from cookie");
        persist_superglobals(&first_vm, &first_ctx);
        assert!(PHP_SESSION_STORE.get(&session_id).is_some());

        let second_ctx = build_ctx(
            "GET",
            "http://localhost:8080/logout.php",
            &[("Host", "localhost:8080"), ("Cookie", &cookie)],
            b"",
        );
        let (_out2, second_ctx, second_vm) = run_php_request_vm(
            r#"<?php session_start(); session_unset(); session_destroy();"#,
            Arc::clone(&second_ctx),
        );
        persist_superglobals(&second_vm, &second_ctx);
        assert!(PHP_SESSION_STORE.get(&session_id).is_none());
        let cleared_cookie = {
            let response = second_ctx.response.lock().expect("lock response");
            response
                .headers
                .iter()
                .find(|(name, _)| name.eq_ignore_ascii_case("set-cookie"))
                .map(|(_, value)| value.clone())
                .expect("cleared set-cookie header")
        };
        let cleared_value = cookie_pair_value(&cleared_cookie, PHP_SESSION_COOKIE_NAME)
            .expect("cleared php session cookie");
        assert!(cleared_value.is_empty());
    }

    #[test]
    fn response_headers_collect_set_cookie_message() {
        let ctx = build_ctx(
            "GET",
            "http://localhost:8080/index.php",
            &[("Host", "localhost:8080")],
            b"",
        );
        let (_out, ctx) = run_php_request(r#"<?php setcookie('a', 'b');"#, ctx);
        let response = ctx.response.lock().expect("lock response");
        assert!(response.headers.iter().any(|(name, value)| {
            name.eq_ignore_ascii_case("set-cookie") && value.starts_with("a=b")
        }));
    }

    #[test]
    fn header_location_defaults_response_to_302() {
        let ctx = build_ctx(
            "GET",
            "http://localhost:8080/index.php",
            &[("Host", "localhost:8080")],
            b"",
        );
        let (_out, ctx) =
            run_php_request(r#"<?php header('Location: /login.php'); echo 'body';"#, ctx);
        let response = ctx.response.lock().expect("lock response");
        assert_eq!(response.status, 302);
        assert!(response.headers.iter().any(|(name, value)| {
            name.eq_ignore_ascii_case("location") && value == "/login.php"
        }));
    }

    #[test]
    fn bare_exit_stops_following_output() {
        let ctx = build_ctx(
            "GET",
            "http://localhost:8080/index.php",
            &[("Host", "localhost:8080")],
            b"",
        );
        let (out, _ctx) = run_php_request(r#"<?php echo 'before'; exit; echo 'after';"#, ctx);
        assert_eq!(out, vec!["before".to_string()]);
    }
}

/// Globals holding this request's `wasi:http` handles.
///
/// The spec has the HOST create `incoming-request` / `response-outparam` and
/// hand them to the guest — for a component that is `incoming-handler.handle`'s
/// two arguments. `vybex --serve` compiles a SCRIPT rather than a component, so
/// there is no export to call: the handles are published as globals under
/// reserved names instead, and the request-shaping primitives read them.
pub const WASI_REQUEST_GLOBAL: &str = "__wasi_http_incoming_request";
pub const WASI_RESPONSE_OUT_GLOBAL: &str = "__wasi_http_response_out";

/// Build the `wasi:http` view of this request and expose it to the VM.
fn install_wasi_http_request(vm: &mut vybe_runtime::VM, ctx: &Arc<RequestContext>) {
    let headers: Vec<(String, Vec<u8>)> = ctx
        .headers
        .iter()
        .map(|(name, value)| (name.clone(), value.as_bytes().to_vec()))
        .collect();

    // Mirror rather than consume: the legacy streaming reader is still the
    // one `node:http.body_read` serves.
    let body = ctx
        .body
        .lock()
        .map(|reader| reader.peek_all().to_vec())
        .unwrap_or_default();

    publish_wasi_request(
        vm,
        &ctx.method,
        &ctx.path,
        &ctx.query,
        &ctx.scheme,
        &ctx.host,
        headers,
        body,
    );
    publish_server_env(vm, ctx);
}

/// Publish the DEPLOYMENT half of the CGI environment.
///
/// `wasi:http` models the message; it has no document root, script path,
/// server identity, peer address or protocol version. Those come from the
/// transport — under their standard CGI names, so the map is language-neutral
/// and `primitives/http_request_env` merges it without renaming anything.
/// The message-derived keys (`REQUEST_METHOD`, `PATH_INFO`, `HTTP_*`, …) are
/// NOT set here: the primitive derives them from `wasi:http` so every language
/// gets them, including ones that never run under this server.
fn publish_server_env(vm: &mut vybe_runtime::VM, ctx: &Arc<RequestContext>) {
    use indexmap::IndexMap;
    use vybe_runtime::Value;
    use vybe_runtime::value::{Object, ObjectKind};

    // No key list: `build_cgi_env` builds the deployment half and NOTHING else,
    // so publishing it whole is publishing exactly those keys. The literal
    // 18-name filter that used to sit here was a second copy of that function's
    // key set, kept in step by hand.
    let mut entries: IndexMap<Value, Value> = IndexMap::new();
    for (key, value) in &ctx.env {
        entries.insert(
            Value::String(std::sync::Arc::from(key.as_str())),
            Value::String(std::sync::Arc::from(value.as_str())),
        );
    }
    // Peer address is the socket's, not the message's.
    entries.insert(
        Value::String(std::sync::Arc::from("REMOTE_ADDR")),
        Value::String(std::sync::Arc::from(ctx.remote_addr.as_str())),
    );
    entries.insert(
        Value::String(std::sync::Arc::from("REMOTE_PORT")),
        Value::String(std::sync::Arc::from(ctx.remote_port.to_string().as_str())),
    );

    let mut object = Object::new();
    object.kind = ObjectKind::Map(entries);
    vm.globals.insert(
        vybe_compiler::primitives::http_request_env::SERVER_ENV_GLOBAL.to_string(),
        Value::Object(vybe_runtime::heap::alloc(object)),
    );
}

/// Map raw request parts onto `wasi:http` handles and publish them as globals.
///
/// Split out from [`install_wasi_http_request`] so the MAPPING is testable
/// without a hyper request: joining path+query, deriving the authority, and
/// treating an empty scheme as absent are all places to get the spec wrong.
#[allow(clippy::too_many_arguments)]
pub fn publish_wasi_request(
    vm: &mut vybe_runtime::VM,
    method: &str,
    path: &str,
    query: &str,
    scheme: &str,
    host: &str,
    headers: Vec<(String, Vec<u8>)>,
    body: Vec<u8>,
) -> (u32, u32) {
    // §incoming-request.path-with-query is the path AND query together.
    let path_with_query = if query.is_empty() {
        path.to_string()
    } else {
        format!("{path}?{query}")
    };

    // `scheme` and `authority` are `option<…>`: empty means absent, not "".
    let scheme = (!scheme.is_empty()).then(|| scheme.to_string());
    let authority = (!host.is_empty()).then(|| host.to_string());

    let request_id = vybe_platform_wasi::http::push_incoming_request(
        method,
        Some(path_with_query),
        scheme,
        authority,
        headers,
        body,
    );
    let param_id = vybe_platform_wasi::http::push_response_outparam();

    if let Some(value) = vybe_platform_wasi::http::incoming_request_value(vm, request_id) {
        vm.globals.insert(WASI_REQUEST_GLOBAL.to_string(), value);
    }
    if let Some(value) = vybe_platform_wasi::http::response_outparam_value(vm, param_id) {
        vm.globals
            .insert(WASI_RESPONSE_OUT_GLOBAL.to_string(), value);
    }
    (request_id, param_id)
}
