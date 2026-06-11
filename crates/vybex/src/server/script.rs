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

use super::response_stream::{BoxBody, build_response, bytes_response};
use bytes::Bytes;
use http::Response;
use indexmap::IndexMap;
use vybe_host::{Capabilities, Capability, RequestContext};

const PHP_SESSION_COOKIE_NAME: &str = "PHPSESSID";
const PHP_SESSION_ID_GLOBAL: &str = "__php_session_id";
const PHP_SESSION_STARTED_GLOBAL: &str = "__php_session_started";
const PHP_SESSION_NEEDS_COOKIE_GLOBAL: &str = "__php_session_needs_cookie";
const PHP_SESSION_DESTROYED_GLOBAL: &str = "__php_session_destroyed";

static PHP_SESSION_STORE: std::sync::LazyLock<
    dashmap::DashMap<String, IndexMap<String, vybe_bytecode::Value>>,
> = std::sync::LazyLock::new(dashmap::DashMap::new);

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
    let bundle = match vybe_compiler::projects::load(script_path) {
        Ok(b) => b,
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
    vybe_host::setup_namespaces(&mut vm);

    crate::server::programmatic::register(&mut vm);
    if let Err(e) = crate::adapters::register_all(&mut vm) {
        let msg = format!("adapter registration error: {e}");
        end_with_text(&ctx, 500, &msg);
        return;
    }

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
    vm.register_host_fn(
        "wasi:logging/logging",
        "log",
        Box::new(|_ctx, args| {
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
        }),
    );

    let mut runtime_compiler = crate::dynamic::RuntimeCompilerService::new(&mut vm);
    if let Err(e) = runtime_compiler.compile_and_run_bundle(&bundle) {
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

    persist_superglobals(&vm, &ctx);

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
    use std::sync::{Arc as StdArc, Mutex as StdMutex};
    use vybe_bytecode::Value;
    use vybe_bytecode::value::{Object, ObjectKind};

    let server = make_string_map_value(ctx.env.iter().map(|(k, v)| (k.clone(), v.clone())));
    let env = make_string_map_value(std::env::vars());

    let get_pairs: Vec<(String, String)> = form_urlencoded::parse(ctx.query.as_bytes())
        .map(|(k, v)| (k.into_owned(), v.into_owned()))
        .collect();
    let get = make_string_map_value(get_pairs);

    let cookie_header = ctx
        .headers
        .iter()
        .find(|(n, _)| n.eq_ignore_ascii_case("cookie"))
        .map(|(_, v)| v.as_str())
        .unwrap_or("");
    let cookie_pairs = parse_cookie_header(cookie_header);
    let session_cookie = cookie_pairs
        .iter()
        .find(|(name, value)| name == PHP_SESSION_COOKIE_NAME && !value.is_empty())
        .map(|(_, value)| value.clone());
    let cookies = make_string_map_value(cookie_pairs);
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

    let content_type = ctx
        .headers
        .iter()
        .find(|(n, _)| n.eq_ignore_ascii_case("content-type"))
        .map(|(_, v)| v.as_str())
        .unwrap_or("");
    let body = ctx.body.lock().unwrap().read_all();
    let (post_pairs, file_pairs) = if content_type
        .to_ascii_lowercase()
        .starts_with("application/x-www-form-urlencoded")
    {
        (
            form_urlencoded::parse(&body)
                .map(|(k, v)| (k.into_owned(), v.into_owned()))
                .collect(),
            Vec::new(),
        )
    } else if content_type
        .to_ascii_lowercase()
        .starts_with("multipart/form-data")
    {
        parse_multipart_form(&body, content_type)
    } else {
        (Vec::new(), Vec::new())
    };
    let post = make_string_map_value(post_pairs);
    let files = make_map_value(file_pairs);

    // PHP variables and functions live in separate namespaces; the
    // walker preserves the `$` sigil on variable identifiers so a
    // function `foo` and a variable `$foo` don't collide. Register
    // the superglobals with the same `$` prefix so user code's
    // `$_SERVER["PHP_SELF"]` etc. resolves correctly.
    vm.globals.insert("$_SERVER".to_string(), server);
    vm.globals.insert("$_ENV".to_string(), env);
    vm.globals.insert("$_GET".to_string(), get);
    vm.globals.insert("$_COOKIE".to_string(), cookies);
    vm.globals.insert("$_POST".to_string(), post);
    vm.globals.insert("$_FILES".to_string(), files);
    vm.globals.insert("$_SESSION".to_string(), session);
    vm.globals.insert(
        PHP_SESSION_ID_GLOBAL.to_string(),
        Value::String(StdArc::from(session_id.as_str())),
    );
    vm.globals
        .insert(PHP_SESSION_STARTED_GLOBAL.to_string(), Value::Bool(false));
    vm.globals.insert(
        PHP_SESSION_NEEDS_COOKIE_GLOBAL.to_string(),
        Value::Bool(session_cookie.is_none()),
    );
    vm.globals
        .insert(PHP_SESSION_DESTROYED_GLOBAL.to_string(), Value::Bool(false));
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

fn make_map_value(
    pairs: impl IntoIterator<Item = (String, vybe_bytecode::Value)>,
) -> vybe_bytecode::Value {
    use indexmap::IndexMap;
    use std::sync::{Arc as StdArc, Mutex as StdMutex};
    use vybe_bytecode::Value;
    use vybe_bytecode::value::{Object, ObjectKind};

    let mut im = IndexMap::new();
    for (k, v) in pairs {
        im.insert(Value::String(StdArc::from(k.as_str())), v);
    }
    let mut obj = Object::new();
    obj.kind = ObjectKind::Map(im);
    Value::Object(StdArc::new(StdMutex::new(obj)))
}

fn make_string_map_value(
    pairs: impl IntoIterator<Item = (String, String)>,
) -> vybe_bytecode::Value {
    make_map_value(pairs.into_iter().map(|(k, v)| {
        (
            k,
            vybe_bytecode::Value::String(std::sync::Arc::from(v.as_str())),
        )
    }))
}

fn persist_superglobals(vm: &vybe_bytecode::VM, _ctx: &Arc<RequestContext>) {
    let destroyed = matches!(
        vm.globals.get(PHP_SESSION_DESTROYED_GLOBAL),
        Some(vybe_bytecode::Value::Bool(true))
    );

    if destroyed {
        let session_id = match vm.globals.get(PHP_SESSION_ID_GLOBAL) {
            Some(vybe_bytecode::Value::String(s)) if !s.is_empty() => s.to_string(),
            _ => return,
        };
        PHP_SESSION_STORE.remove(&session_id);
        return;
    }

    let started = matches!(
        vm.globals.get(PHP_SESSION_STARTED_GLOBAL),
        Some(vybe_bytecode::Value::Bool(true))
    );
    if !started {
        return;
    }

    let session_id = match vm.globals.get(PHP_SESSION_ID_GLOBAL) {
        Some(vybe_bytecode::Value::String(s)) if !s.is_empty() => s.to_string(),
        _ => return,
    };

    let Some(vybe_bytecode::Value::Object(obj)) = vm.globals.get("$_SESSION") else {
        return;
    };
    let guard = obj.lock().unwrap();
    let vybe_bytecode::value::ObjectKind::Map(map) = &guard.kind else {
        return;
    };

    let persisted: IndexMap<String, vybe_bytecode::Value> = map
        .iter()
        .map(|(key, value)| (format!("{}", key), value.clone()))
        .collect();
    PHP_SESSION_STORE.insert(session_id, persisted);
}

fn parse_cookie_header(header: &str) -> Vec<(String, String)> {
    let mut out = Vec::new();
    for part in header.split(';') {
        let part = part.trim();
        if part.is_empty() {
            continue;
        }
        match part.split_once('=') {
            Some((n, v)) => {
                let v = v
                    .strip_prefix('"')
                    .and_then(|s| s.strip_suffix('"'))
                    .unwrap_or(v);
                let decoded = percent_encoding::percent_decode_str(v)
                    .decode_utf8_lossy()
                    .into_owned();
                out.push((n.trim().to_string(), decoded));
            }
            None => out.push((part.to_string(), String::new())),
        }
    }
    out
}

fn parse_multipart_form(
    body: &[u8],
    content_type: &str,
) -> (Vec<(String, String)>, Vec<(String, vybe_bytecode::Value)>) {
    use vybe_bytecode::Value;

    let Some(boundary) = extract_multipart_boundary(content_type) else {
        return (Vec::new(), Vec::new());
    };

    let marker = format!("--{boundary}").into_bytes();
    let mut post = Vec::new();
    let mut files = Vec::new();
    let mut pos = match find_subslice(body, &marker, 0) {
        Some(pos) => pos,
        None => return (post, files),
    };

    loop {
        let mut cursor = pos + marker.len();
        if body.get(cursor..cursor + 2) == Some(b"--") {
            break;
        }
        if body.get(cursor..cursor + 2) == Some(b"\r\n") {
            cursor += 2;
        }

        let Some(next) = find_subslice(body, &marker, cursor) else {
            break;
        };
        let mut part = &body[cursor..next];
        if part.ends_with(b"\r\n") {
            part = &part[..part.len() - 2];
        }
        pos = next;
        if part.is_empty() {
            continue;
        }

        let Some(header_end) = find_subslice(part, b"\r\n\r\n", 0) else {
            continue;
        };
        let header_text = String::from_utf8_lossy(&part[..header_end]);
        let data = &part[header_end + 4..];

        let mut field_name = None;
        let mut file_name = None;
        let mut file_type = String::new();
        for line in header_text.split("\r\n") {
            let lower = line.to_ascii_lowercase();
            if lower.starts_with("content-disposition:") {
                field_name = extract_disposition_param(line, "name");
                file_name = extract_disposition_param(line, "filename");
            } else if lower.starts_with("content-type:") {
                file_type = line
                    .split_once(':')
                    .map(|(_, value)| value.trim().to_string())
                    .unwrap_or_default();
            }
        }

        let Some(name) = field_name else {
            continue;
        };
        if let Some(filename) = file_name {
            let (tmp_name, error_code) = write_upload_tempfile(data);
            let upload = make_map_value(vec![
                (
                    "name".to_string(),
                    Value::String(std::sync::Arc::from(filename.as_str())),
                ),
                (
                    "type".to_string(),
                    Value::String(std::sync::Arc::from(file_type.as_str())),
                ),
                (
                    "tmp_name".to_string(),
                    Value::String(std::sync::Arc::from(tmp_name.as_str())),
                ),
                ("error".to_string(), Value::F64(error_code as f64)),
                ("size".to_string(), Value::F64(data.len() as f64)),
            ]);
            files.push((name, upload));
        } else {
            post.push((name, String::from_utf8_lossy(data).into_owned()));
        }
    }

    (post, files)
}

fn extract_multipart_boundary(content_type: &str) -> Option<String> {
    content_type
        .split(';')
        .map(str::trim)
        .find_map(|part| part.strip_prefix("boundary="))
        .map(|value| value.trim_matches('"').to_string())
}

fn extract_disposition_param(line: &str, key: &str) -> Option<String> {
    let needle = format!("{key}=\"");
    let start = line.find(&needle)? + needle.len();
    let rest = &line[start..];
    let end = rest.find('"')?;
    Some(rest[..end].to_string())
}

fn write_upload_tempfile(data: &[u8]) -> (String, u32) {
    let path = std::env::temp_dir().join(format!("vybex-upload-{}", uuid::Uuid::new_v4()));
    match std::fs::write(&path, data) {
        Ok(()) => (path.display().to_string(), 0),
        Err(_) => (String::new(), 7),
    }
}

fn find_subslice(haystack: &[u8], needle: &[u8], start: usize) -> Option<usize> {
    if needle.is_empty() || start >= haystack.len() {
        return None;
    }
    haystack[start..]
        .windows(needle.len())
        .position(|window| window == needle)
        .map(|offset| start + offset)
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
        PHP_SESSION_STORE, inject_superglobals, parse_cookie_header, persist_superglobals,
    };
    use bytes::Bytes;
    use http::Request;
    use indexmap::IndexMap;
    use std::net::{Ipv4Addr, SocketAddr, SocketAddrV4};
    use std::path::Path;
    use std::sync::{Arc, Mutex};
    use vybe_bytecode::value::ObjectKind;
    use vybe_bytecode::{HostContext, VM, Value};
    fn compile_php(src: &str) -> Vec<vybe_bytecode::Chunk> {
        let module = vybe_compiler::languages::php::parse(src).expect("parse php");
        let profile =
            vybe_compiler::profile::parse_profile(vybe_compiler::languages::php::profile_source())
                .expect("parse php profile");
        vybe_compiler::compiler::Compiler::with_profile(profile)
            .compile(&module)
            .expect("compile php")
    }

    fn build_ctx(
        method: &str,
        uri: &str,
        headers: &[(&str, &str)],
        body: &[u8],
    ) -> Arc<vybe_host::RequestContext> {
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

    fn map_value<'a>(map: &'a IndexMap<String, Value>, key: &str) -> &'a Value {
        map.get(key).unwrap_or_else(|| panic!("missing key {key}"))
    }

    fn value_as_string(value: &Value) -> String {
        match value {
            Value::String(s) => s.to_string(),
            other => format!("{}", other),
        }
    }

    fn run_php_request_vm(
        src: &str,
        ctx: Arc<vybe_host::RequestContext>,
    ) -> (Vec<String>, Arc<vybe_host::RequestContext>, VM) {
        let chunks = compile_php(src);
        let mut vm = VM::new();
        let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
        let out = Arc::clone(&output);
        vybe_host::register_all(&mut vm);
        vm.register_host_fn(
            "wasi:logging/logging",
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
        vybe_host::setup_namespaces(&mut vm);
        inject_superglobals(&mut vm, &ctx);
        let _guard = vybe_host::install_context(Arc::clone(&ctx));
        vm.run(chunks).expect("run php request");
        (output.lock().expect("lock output").clone(), ctx, vm)
    }

    fn run_php_request(
        src: &str,
        ctx: Arc<vybe_host::RequestContext>,
    ) -> (Vec<String>, Arc<vybe_host::RequestContext>) {
        let (output, ctx, vm) = run_php_request_vm(src, ctx);
        persist_superglobals(&vm, &ctx);
        (output, ctx)
    }

    #[test]
    fn inject_superglobals_adds_server_timing_and_env_keys() {
        let ctx = build_ctx(
            "GET",
            "http://localhost:8080/index.php?foo=bar",
            &[("Host", "localhost:8080")],
            b"",
        );
        let mut vm = VM::new();
        inject_superglobals(&mut vm, &ctx);

        let server = map_entries(vm.globals.get("$_SERVER").expect("$_SERVER"));
        assert_eq!(
            value_as_string(map_value(&server, "PHP_SELF")),
            "/index.php"
        );
        assert_eq!(
            value_as_string(map_value(&server, "SCRIPT_NAME")),
            "/index.php"
        );
        assert_eq!(
            value_as_string(map_value(&server, "SCRIPT_FILENAME")),
            "/tmp/index.php"
        );
        assert_eq!(
            value_as_string(map_value(&server, "PATH_TRANSLATED")),
            "/tmp/index.php"
        );
        assert_eq!(
            value_as_string(map_value(&server, "DOCUMENT_URI")),
            "/index.php"
        );
        assert_eq!(
            value_as_string(map_value(&server, "SCRIPT_URL")),
            "/index.php"
        );
        assert_eq!(
            value_as_string(map_value(&server, "SCRIPT_URI")),
            "http://localhost:8080/index.php?foo=bar"
        );
        assert_eq!(
            value_as_string(map_value(&server, "HTTP_HOST")),
            "localhost:8080"
        );
        assert_eq!(
            value_as_string(map_value(&server, "REQUEST_SCHEME")),
            "http"
        );
        assert_eq!(
            value_as_string(map_value(&server, "SERVER_ADDR")),
            "127.0.0.1"
        );
        assert_eq!(
            value_as_string(map_value(&server, "REMOTE_HOST")),
            "127.0.0.1"
        );
        assert!(!value_as_string(map_value(&server, "REQUEST_TIME")).is_empty());
        assert!(!value_as_string(map_value(&server, "REQUEST_TIME_FLOAT")).is_empty());

        let env = map_entries(vm.globals.get("$_ENV").expect("$_ENV"));
        assert!(env.contains_key("HOME") || env.contains_key("PATH"));
    }

    #[test]
    fn inject_superglobals_uses_resolved_script_name_for_directory_indexes() {
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

        let mut vm = VM::new();
        inject_superglobals(&mut vm, &built.ctx);

        let server = map_entries(vm.globals.get("$_SERVER").expect("$_SERVER"));
        assert_eq!(
            value_as_string(map_value(&server, "REQUEST_URI")),
            "http://localhost:8080/genie/"
        );
        assert_eq!(
            value_as_string(map_value(&server, "SCRIPT_NAME")),
            "/genie/index.php"
        );
        assert_eq!(
            value_as_string(map_value(&server, "PHP_SELF")),
            "/genie/index.php"
        );
    }

    #[test]
    fn inject_superglobals_parses_multipart_into_post_and_files() {
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
        let mut vm = VM::new();
        inject_superglobals(&mut vm, &ctx);

        let post = map_entries(vm.globals.get("$_POST").expect("$_POST"));
        assert_eq!(value_as_string(map_value(&post, "title")), "hello");

        let files = map_entries(vm.globals.get("$_FILES").expect("$_FILES"));
        let upload = map_entries(map_value(&files, "upload"));
        assert_eq!(value_as_string(map_value(&upload, "name")), "note.txt");
        assert_eq!(value_as_string(map_value(&upload, "type")), "text/plain");
        assert_eq!(value_as_string(map_value(&upload, "size")), "5");
        let tmp_name = value_as_string(map_value(&upload, "tmp_name"));
        assert!(!tmp_name.is_empty(), "tmp_name should not be empty");
        assert!(
            std::fs::metadata(tmp_name).is_ok(),
            "uploaded tmp file should exist"
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
        let session_id = parse_cookie_header(&cookie)
            .into_iter()
            .find(|(name, _)| name == PHP_SESSION_COOKIE_NAME)
            .map(|(_, value)| value)
            .expect("session id from cookie");
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
        let session_id = parse_cookie_header(&cookie)
            .into_iter()
            .find(|(name, _)| name == PHP_SESSION_COOKIE_NAME)
            .map(|(_, value)| value)
            .expect("session id from cookie");
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
        let cleared_value = parse_cookie_header(&cleared_cookie)
            .into_iter()
            .find(|(name, _)| name == PHP_SESSION_COOKIE_NAME)
            .map(|(_, value)| value)
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
