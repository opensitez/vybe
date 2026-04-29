//! `node:http.*` — per-request response write.
//!
//! All mutators silently no-op when no request context is installed.
//! Setting headers / status after the first body write is also a no-op
//! (matching PHP's `headers_sent` warning behavior).

use std::sync::Arc;
use vybe_bytecode::{VM, Value, HostContext};

use super::context::with_context;

pub fn register(vm: &mut VM) {
    // Status ───────────────────────────────────────────────────────────────
    vm.register_host_fn("node:http", "set_status", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let code = args.first().map(|v| v.as_f64() as u16).unwrap_or(200);
        with_context(|c| {
            let mut r = c.response.lock().unwrap();
            if !r.headers_sent { r.status = code; }
        });
        Value::Null
    }));

    vm.register_host_fn("node:http", "status", Box::new(|_ctx, _| {
        with_context(|c| Value::F64(c.response.lock().unwrap().status as f64))
            .unwrap_or(Value::F64(0.0))
    }));

    // Headers ──────────────────────────────────────────────────────────────
    vm.register_host_fn("node:http", "set_header", Box::new(|_ctx, args| {
        let name = string_arg(args, 0);
        let value = string_arg(args, 1);
        with_context(|c| {
            let mut r = c.response.lock().unwrap();
            if r.headers_sent { return; }
            r.headers.retain(|(n, _)| !n.eq_ignore_ascii_case(&name));
            r.headers.push((name, value));
        });
        Value::Null
    }));

    vm.register_host_fn("node:http", "add_header", Box::new(|_ctx, args| {
        let name = string_arg(args, 0);
        let value = string_arg(args, 1);
        with_context(|c| {
            let mut r = c.response.lock().unwrap();
            if r.headers_sent { return; }
            r.headers.push((name, value));
        });
        Value::Null
    }));

    vm.register_host_fn("node:http", "remove_header", Box::new(|_ctx, args| {
        let name = string_arg(args, 0);
        with_context(|c| {
            let mut r = c.response.lock().unwrap();
            if r.headers_sent { return; }
            r.headers.retain(|(n, _)| !n.eq_ignore_ascii_case(&name));
        });
        Value::Null
    }));

    vm.register_host_fn("node:http", "has_header", Box::new(|_ctx, args| {
        let name = string_arg(args, 0);
        with_context(|c| {
            let r = c.response.lock().unwrap();
            Value::Bool(r.headers.iter().any(|(n, _)| n.eq_ignore_ascii_case(&name)))
        }).unwrap_or(Value::Bool(false))
    }));

    vm.register_host_fn("node:http", "headers_sent", Box::new(|_ctx, _| {
        with_context(|c| Value::Bool(c.response.lock().unwrap().headers_sent))
            .unwrap_or(Value::Bool(false))
    }));

    // Body ─────────────────────────────────────────────────────────────────
    vm.register_host_fn("node:http", "write", Box::new(|_ctx, args| {
        let bytes = match args.first() {
            Some(Value::String(s)) => s.as_bytes().to_vec(),
            Some(other) => format!("{}", other).into_bytes(),
            None => Vec::new(),
        };
        with_context(|c| {
            c.response.lock().unwrap().write_bytes(bytes);
        });
        Value::Null
    }));

    vm.register_host_fn("node:http", "write_text", Box::new(|_ctx, args| {
        let text = string_arg(args, 0);
        with_context(|c| {
            c.response.lock().unwrap().write_bytes(text.into_bytes());
        });
        Value::Null
    }));

    vm.register_host_fn("node:http", "end", Box::new(|_ctx, _| {
        with_context(|c| {
            c.response.lock().unwrap().end();
        });
        Value::Null
    }));

    vm.register_host_fn("node:http", "flush", Box::new(|_ctx, _| {
        // Phase 1: writes are unbuffered on the host side (each write_bytes
        // pushes a Data message). Explicit flush is a no-op for now.
        // Phase 2: add an ~8 KiB coalescing buffer and wire flush() to push
        // its contents.
        Value::Null
    }));

    // ── http_response_code ─────────────────────────────────────────────────
    //
    // PHP idiom: `http_response_code()` returns the current status;
    // `http_response_code(404)` sets it. Combined getter/setter in one
    // host fn since PHP's dispatch is arity-based.
    vm.register_host_fn("node:http", "http_response_code", Box::new(|_ctx, args| {
        if let Some(arg) = args.first() {
            let code = arg.as_f64() as u16;
            if code > 0 {
                with_context(|c| {
                    let mut r = c.response.lock().unwrap();
                    if !r.headers_sent { r.status = code; }
                });
                return Value::F64(code as f64);
            }
        }
        with_context(|c| Value::F64(c.response.lock().unwrap().status as f64))
            .unwrap_or(Value::F64(200.0))
    }));

    // ── send_header_raw ────────────────────────────────────────────────────
    //
    // PHP `header()` idiom: accept a raw header line
    // (`"Content-Type: text/plain"` or `"HTTP/1.1 404 Not Found"`) plus an
    // optional `replace` flag and optional status code. Parse it centrally,
    // route to set_status / set_header / add_header. One implementation,
    // every language that has a PHP-compat wrapper benefits.
    vm.register_host_fn("node:http", "send_header_raw", Box::new(|_ctx, args| {
        let raw = string_arg(args, 0);
        let replace = args.get(1).map(|v| v.as_bool()).unwrap_or(true);
        let response_code = args.get(2).map(|v| v.as_f64() as u16).unwrap_or(0);

        with_context(|c| {
            let mut r = c.response.lock().unwrap();
            if r.headers_sent { return; }

            // Status-line form: "HTTP/1.1 404 Not Found"
            if raw.starts_with("HTTP/") {
                let parts: Vec<&str> = raw.splitn(3, ' ').collect();
                if let Some(code) = parts.get(1).and_then(|c| c.parse::<u16>().ok()) {
                    r.status = code;
                }
                return;
            }

            if let Some((name, value)) = raw.split_once(':') {
                let name = name.trim().to_string();
                let value = value.trim().to_string();

                // Special case: `Status: 404 Not Found` maps to set_status.
                if name.eq_ignore_ascii_case("status") {
                    if let Some(code) = value.split_whitespace().next()
                        .and_then(|c| c.parse::<u16>().ok()) {
                        r.status = code;
                    }
                    return;
                }

                if replace {
                    r.headers.retain(|(n, _)| !n.eq_ignore_ascii_case(&name));
                }
                r.headers.push((name, value));
            }

            if response_code > 0 {
                r.status = response_code;
            }
        });
        Value::Null
    }));

    // ── set_cookie ─────────────────────────────────────────────────────────
    //
    // PHP `setcookie(name, value, options)` — serialize a Set-Cookie header
    // from structured args. Options: an associative Map with keys
    // `expires` (int unix), `path`, `domain`, `secure`, `httponly`,
    // `samesite`. Called by PHP's stdlib `setcookie()` idiom.
    vm.register_host_fn("node:http", "set_cookie", Box::new(|_ctx, args| {
        let name = string_arg(args, 0);
        let value = string_arg(args, 1);
        if name.is_empty() { return Value::Bool(false); }

        // Build the Set-Cookie header value.
        let mut out = format!("{}={}", name, url_encode_cookie_value(&value));

        if let Some(Value::Object(opts_obj)) = args.get(2) {
            let opts = opts_obj.lock().unwrap();
            // Options may arrive as either a Map or an Ordinary object
            // (PHP stdlib construction varies). Read both shapes.
            let get = |key: &str| -> Option<String> {
                if let vybe_bytecode::value::ObjectKind::Map(m) = &opts.kind {
                    let as_str = m.get(&Value::String(std::sync::Arc::from(key)));
                    return as_str.map(|v| format!("{}", v));
                }
                opts.properties.get(key).map(|v| format!("{}", v))
            };
            if let Some(v) = get("expires") {
                if let Ok(ts) = v.parse::<i64>() {
                    if ts > 0 {
                        let http_date = httpdate::fmt_http_date(
                            std::time::UNIX_EPOCH + std::time::Duration::from_secs(ts as u64)
                        );
                        out.push_str("; Expires=");
                        out.push_str(&http_date);
                    }
                }
            }
            if let Some(v) = get("path") { out.push_str("; Path="); out.push_str(&v); }
            if let Some(v) = get("domain") { out.push_str("; Domain="); out.push_str(&v); }
            if let Some(v) = get("samesite") { out.push_str("; SameSite="); out.push_str(&v); }
            if get("secure").filter(|v| v == "true" || v == "1").is_some() { out.push_str("; Secure"); }
            if get("httponly").filter(|v| v == "true" || v == "1").is_some() { out.push_str("; HttpOnly"); }
        }

        with_context(|c| {
            let mut r = c.response.lock().unwrap();
            if r.headers_sent { return; }
            r.headers.push(("Set-Cookie".to_string(), out));
        });
        Value::Bool(true)
    }));
}

fn url_encode_cookie_value(v: &str) -> String {
    // PHP's setcookie URL-encodes the value (setrawcookie doesn't). This
    // is a minimal encoder — spaces → %20, control chars → percent, etc.
    // Using percent-encoding for consistency with every other URL path.
    percent_encoding::utf8_percent_encode(v, percent_encoding::NON_ALPHANUMERIC).to_string()
}

fn string_arg(args: &[Value], idx: usize) -> String {
    match args.get(idx) {
        Some(Value::String(s)) => s.to_string(),
        Some(other) => format!("{}", other),
        None => String::new(),
    }
}

// Arc<str> is Send+Sync, so we're good for the closures above.
// Re-export to satisfy unused import warnings if needed.
#[allow(dead_code)]
fn _ensure_arc_in_scope(s: &str) -> Arc<str> { Arc::from(s) }
