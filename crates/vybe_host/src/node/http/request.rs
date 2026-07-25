//! `node:http.*` — per-request read access.
//!
//! All accessors return sentinel values (empty string, null, 0, empty
//! array) when no request context is installed on the current thread
//! ("cli" mode). This makes the primitives safe to call from any script,
//! whether it runs under `--serve` or at the terminal.
//!
//! Parsed accessors (cookies, query pairs, content-type, form bodies,
//! auth) will land in a subsequent slice — they cache their result on the
//! `RequestContext` via `OnceLock`.

use std::sync::Arc;
use vybe_bytecode::value::{Object, ObjectKind};
use vybe_bytecode::{HostContext, VM, Value};

use super::context::with_context;

pub fn register(vm: &mut VM) {
    // Raw accessors ────────────────────────────────────────────────────────
    vm.register_host_fn(
        "node:http",
        "method",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            with_context(|c| Value::String(Arc::from(c.method.as_str())))
                .unwrap_or_else(|| Value::String(Arc::from("")))
        }),
    );

    vm.register_host_fn(
        "node:http",
        "uri",
        Box::new(|_ctx, _| {
            with_context(|c| Value::String(Arc::from(c.uri.as_str())))
                .unwrap_or_else(|| Value::String(Arc::from("")))
        }),
    );

    vm.register_host_fn(
        "node:http",
        "path",
        Box::new(|_ctx, _| {
            with_context(|c| Value::String(Arc::from(c.path.as_str())))
                .unwrap_or_else(|| Value::String(Arc::from("")))
        }),
    );

    vm.register_host_fn(
        "node:http",
        "query",
        Box::new(|_ctx, _| {
            with_context(|c| Value::String(Arc::from(c.query.as_str())))
                .unwrap_or_else(|| Value::String(Arc::from("")))
        }),
    );

    vm.register_host_fn(
        "node:http",
        "scheme",
        Box::new(|_ctx, _| {
            with_context(|c| Value::String(Arc::from(c.scheme.as_str())))
                .unwrap_or_else(|| Value::String(Arc::from("")))
        }),
    );

    vm.register_host_fn(
        "node:http",
        "host",
        Box::new(|_ctx, _| {
            with_context(|c| Value::String(Arc::from(c.host.as_str())))
                .unwrap_or_else(|| Value::String(Arc::from("")))
        }),
    );

    vm.register_host_fn(
        "node:http",
        "port",
        Box::new(|_ctx, _| with_context(|c| Value::F64(c.port as f64)).unwrap_or(Value::F64(0.0))),
    );

    vm.register_host_fn(
        "node:http",
        "protocol",
        Box::new(|_ctx, _| {
            with_context(|c| Value::String(Arc::from(c.protocol.as_str())))
                .unwrap_or_else(|| Value::String(Arc::from("")))
        }),
    );

    vm.register_host_fn(
        "node:http",
        "remote_addr",
        Box::new(|_ctx, _| {
            with_context(|c| Value::String(Arc::from(c.remote_addr.as_str())))
                .unwrap_or_else(|| Value::String(Arc::from("")))
        }),
    );

    vm.register_host_fn(
        "node:http",
        "remote_port",
        Box::new(|_ctx, _| {
            with_context(|c| Value::F64(c.remote_port as f64)).unwrap_or(Value::F64(0.0))
        }),
    );

    // Headers ──────────────────────────────────────────────────────────────
    vm.register_host_fn(
        "node:http",
        "header",
        Box::new(|_ctx, args| {
            let name = string_arg(args, 0).to_ascii_lowercase();
            with_context(|c| {
                for (n, v) in &c.headers {
                    if n.eq_ignore_ascii_case(&name) {
                        return Value::String(Arc::from(v.as_str()));
                    }
                }
                Value::Null
            })
            .unwrap_or(Value::Null)
        }),
    );

    vm.register_host_fn(
        "node:http",
        "header_all",
        Box::new(|_ctx, args| {
            let name = string_arg(args, 0).to_ascii_lowercase();
            with_context(|c| {
                let mut out = Vec::new();
                for (n, v) in &c.headers {
                    if n.eq_ignore_ascii_case(&name) {
                        out.push(Value::String(Arc::from(v.as_str())));
                    }
                }
                array_value(out)
            })
            .unwrap_or_else(|| array_value(Vec::new()))
        }),
    );

    vm.register_host_fn(
        "node:http",
        "headers",
        Box::new(|_ctx, _| {
            with_context(|c| {
                let items = c.headers.iter().map(|(n, v)| pair_object(n, v)).collect();
                array_value(items)
            })
            .unwrap_or_else(|| array_value(Vec::new()))
        }),
    );

    // CGI env ──────────────────────────────────────────────────────────────
    vm.register_host_fn(
        "node:http",
        "env",
        Box::new(|_ctx, args| {
            let name = string_arg(args, 0);
            with_context(|c| match c.env.get(&name) {
                Some(v) => Value::String(Arc::from(v.as_str())),
                None => Value::Null,
            })
            .unwrap_or(Value::Null)
        }),
    );

    vm.register_host_fn(
        "node:http",
        "envs",
        Box::new(|_ctx, _| {
            with_context(|c| {
                let items: Vec<Value> = c.env.iter().map(|(k, v)| pair_object(k, v)).collect();
                array_value(items)
            })
            .unwrap_or_else(|| array_value(Vec::new()))
        }),
    );

    // Body streaming ───────────────────────────────────────────────────────
    vm.register_host_fn(
        "node:http",
        "body_length",
        Box::new(|_ctx, _| {
            with_context(|c| match c.body.lock().unwrap().length() {
                Some(n) => Value::F64(n as f64),
                None => Value::Null,
            })
            .unwrap_or(Value::Null)
        }),
    );

    vm.register_host_fn(
        "node:http",
        "body_eof",
        Box::new(|_ctx, _| {
            with_context(|c| Value::Bool(c.body.lock().unwrap().eof())).unwrap_or(Value::Bool(true))
        }),
    );

    vm.register_host_fn(
        "node:http",
        "body_read",
        Box::new(|_ctx, args| {
            let max = args.first().map(|v| v.as_f64() as usize).unwrap_or(0);
            with_context(|c| {
                let bytes = c.body.lock().unwrap().read(max);
                // Phase 1: return as a string. Binary-correctness upgrade
                // (ArrayBuffer / bytes value) tracked for Phase 2 once we
                // commit to a bytes type.
                Value::String(Arc::from(String::from_utf8_lossy(&bytes).as_ref()))
            })
            .unwrap_or_else(|| Value::String(Arc::from("")))
        }),
    );

    vm.register_host_fn(
        "node:http",
        "body_read_all",
        Box::new(|_ctx, _| {
            with_context(|c| {
                let bytes = c.body.lock().unwrap().read_all();
                Value::String(Arc::from(String::from_utf8_lossy(&bytes).as_ref()))
            })
            .unwrap_or_else(|| Value::String(Arc::from("")))
        }),
    );

    vm.register_host_fn(
        "node:http",
        "request_id",
        Box::new(|_ctx, _| {
            with_context(|c| Value::String(Arc::from(c.request_id.as_str())))
                .unwrap_or_else(|| Value::String(Arc::from("")))
        }),
    );

    // Parsed accessors (centralized — every language's adapter uses these) ─
    vm.register_host_fn(
        "node:http",
        "cookies",
        Box::new(|_ctx, _| {
            with_context(|c| {
                let parsed = c
                    .cookies
                    .get_or_init(|| parse_cookies_from_headers(&c.headers));
                let items: Vec<Value> = parsed.iter().map(|(n, v)| pair_object(n, v)).collect();
                array_value(items)
            })
            .unwrap_or_else(|| array_value(Vec::new()))
        }),
    );

    vm.register_host_fn(
        "node:http",
        "query_pairs",
        Box::new(|_ctx, _| {
            with_context(|c| {
                let parsed = c.query_pairs.get_or_init(|| parse_query(&c.query));
                let items: Vec<Value> = parsed.iter().map(|(n, v)| pair_object(n, v)).collect();
                array_value(items)
            })
            .unwrap_or_else(|| array_value(Vec::new()))
        }),
    );
}

// Parsing helpers — the ONE implementation. Every language's stdlib calls
// the host fn instead of rolling its own parser.
fn parse_cookies_from_headers(headers: &[(String, String)]) -> Vec<(String, String)> {
    let mut out = Vec::new();
    for (n, v) in headers {
        if !n.eq_ignore_ascii_case("cookie") {
            continue;
        }
        // Multiple Cookie headers per RFC 6265 are joined with `; `.
        for part in v.split(';') {
            let part = part.trim();
            if part.is_empty() {
                continue;
            }
            match part.split_once('=') {
                Some((name, value)) => {
                    // Cookie values may be quoted; strip one layer.
                    let value = value.trim();
                    let value = value
                        .strip_prefix('"')
                        .and_then(|s| s.strip_suffix('"'))
                        .unwrap_or(value);
                    out.push((name.trim().to_string(), value.to_string()));
                }
                None => {
                    out.push((part.to_string(), String::new()));
                }
            }
        }
    }
    out
}

fn parse_query(query: &str) -> Vec<(String, String)> {
    form_urlencoded::parse(query.as_bytes())
        .map(|(k, v)| (k.into_owned(), v.into_owned()))
        .collect()
}

// Helpers ──────────────────────────────────────────────────────────────────

fn string_arg(args: &[Value], idx: usize) -> String {
    args.get(idx).map(|v| format!("{}", v)).unwrap_or_default()
}

fn array_value(items: Vec<Value>) -> Value {
    let mut obj = Object::new();
    obj.kind = ObjectKind::Array(items);
    Value::Object(vybe_bytecode::heap::alloc(obj))
}

fn pair_object(name: &str, value: &str) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("name".into(), Value::String(Arc::from(name)));
    obj.properties
        .insert("value".into(), Value::String(Arc::from(value)));
    Value::Object(vybe_bytecode::heap::alloc(obj))
}
