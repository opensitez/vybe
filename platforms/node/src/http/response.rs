//! `node:http.*` — per-request response write.
//!
//! All mutators silently no-op when no request context is installed.
//! Setting headers / status after the first body write is also a no-op
//! (matching PHP's `headers_sent` warning behavior).

use std::sync::Arc;
use vybe_runtime::{HostContext, VM, Value};

use super::context::with_context;

pub fn register(vm: &mut VM) {
    // `ServerResponse.getHeaders()` — a shallow copy of the outgoing headers,
    // keyed by LOWER-CASE name. Node returns a null-prototype object.
    vm.register_host_fn(
        "node:http",
        "get_headers",
        Box::new(|_ctx, _args| {
            let mut entries = indexmap::IndexMap::new();
            with_context(|c| {
                let r = c.response.lock().unwrap();
                for (name, value) in r.headers.iter() {
                    entries.insert(
                        Value::String(Arc::from(name.to_ascii_lowercase().as_str())),
                        Value::String(Arc::from(value.as_str())),
                    );
                }
            });
            let mut object = vybe_runtime::value::Object::new();
            object.kind = vybe_runtime::value::ObjectKind::Map(entries);
            Value::Object(vybe_runtime::heap::alloc(object))
        }),
    );

    // `ServerResponse.addTrailers(headers)` — fields sent after the body.
    // Node only emits them for a chunked response that announced a `Trailer`
    // header, so they are kept apart from the ordinary header list.
    vm.register_host_fn(
        "node:http",
        "add_trailers",
        Box::new(|_ctx, args| {
            let pairs = header_pairs(args.first());
            with_context(|c| {
                let mut r = c.response.lock().unwrap();
                for (name, value) in pairs {
                    r.trailers.push((name, value));
                }
            });
            Value::Null
        }),
    );

    // `ServerResponse.writeContinue()` — the 100 Continue interim response.
    // Interim responses do NOT end the exchange: the final status still
    // follows, so this must not touch `status` or mark headers sent.
    vm.register_host_fn(
        "node:http",
        "write_continue",
        Box::new(|_ctx, _args| {
            with_context(|c| {
                c.response.lock().unwrap().interim.push((100, Vec::new()));
            });
            Value::Null
        }),
    );

    // `ServerResponse.writeEarlyHints(hints)` — a 103 interim response,
    // typically carrying `Link` preload hints (RFC 8297).
    vm.register_host_fn(
        "node:http",
        "write_early_hints",
        Box::new(|_ctx, args| {
            let pairs = header_pairs(args.first());
            with_context(|c| {
                c.response.lock().unwrap().interim.push((103, pairs));
            });
            Value::Null
        }),
    );

    // ── ServerResponse.getHeader / getHeaders / getHeaderNames ─────────────
    //
    // Node returns `undefined` for a header that was never set, and matches
    // names case-INSENSITIVELY while reporting them lower-cased
    // (`ServerResponse.getHeaders`).
    vm.register_host_fn(
        "node:http",
        "get_header",
        Box::new(|_ctx, args| {
            let name = string_arg(args, 0);
            with_context(|c| {
                let r = c.response.lock().unwrap();
                r.headers
                    .iter()
                    .find(|(n, _)| n.eq_ignore_ascii_case(&name))
                    .map(|(_, v)| Value::String(Arc::from(v.as_str())))
                    .unwrap_or(Value::Undefined)
            })
            .unwrap_or(Value::Undefined)
        }),
    );

    vm.register_host_fn(
        "node:http",
        "get_header_names",
        Box::new(|_ctx, _args| {
            let names = with_context(|c| {
                let r = c.response.lock().unwrap();
                let mut seen: Vec<String> = Vec::new();
                for (name, _) in r.headers.iter() {
                    let lower = name.to_ascii_lowercase();
                    if !seen.contains(&lower) {
                        seen.push(lower);
                    }
                }
                seen
            })
            .unwrap_or_default();
            let elems: Vec<Value> = names
                .into_iter()
                .map(|n| Value::String(Arc::from(n.as_str())))
                .collect();
            Value::Object(vybe_runtime::heap::alloc(
                vybe_runtime::value::Object::new_array(elems),
            ))
        }),
    );

    // ── ServerResponse.writeHead(statusCode[, statusMessage][, headers]) ───
    //
    // Node's most-used response call, and it is NOT setHeader in a loop: it
    // sets the status and the headers together, and the headers argument may
    // be an object or a flat `[k, v, k, v]` array.
    vm.register_host_fn(
        "node:http",
        "write_head",
        Box::new(|_ctx, args| {
            let status = args.first().map(|v| v.as_f64() as u16).unwrap_or(200);
            // The optional statusMessage sits between status and headers.
            let (message, headers) = match (args.get(1), args.get(2)) {
                (Some(Value::String(m)), rest) => (Some(m.to_string()), rest),
                (rest, _) => (None, rest),
            };
            with_context(|c| {
                let mut r = c.response.lock().unwrap();
                if r.headers_sent {
                    return;
                }
                if status > 0 {
                    r.status = status;
                }
                if let Some(m) = message {
                    r.status_message = Some(m);
                }
                for (name, value) in header_pairs(headers) {
                    r.headers.retain(|(n, _)| !n.eq_ignore_ascii_case(&name));
                    r.headers.push((name, value));
                }
            });
            Value::Null
        }),
    );

    // `ServerResponse.statusMessage` — the reason phrase. Empty when unset;
    // Node fills it from `STATUS_CODES` at send time.
    vm.register_host_fn(
        "node:http",
        "status_message",
        Box::new(|_ctx, _args| {
            with_context(|c| {
                let r = c.response.lock().unwrap();
                // Unset falls back to the registered reason phrase, which is
                // what Node puts on the wire (`STATUS_CODES[statusCode]`).
                match &r.status_message {
                    Some(m) => Value::String(Arc::from(m.as_str())),
                    None => Value::String(Arc::from(
                        super::tables::reason_phrase(r.status).unwrap_or(""),
                    )),
                }
            })
            .unwrap_or_else(|| Value::String(Arc::from("")))
        }),
    );

    vm.register_host_fn(
        "node:http",
        "set_status_message",
        Box::new(|_ctx, args| {
            let message = string_arg(args, 0);
            with_context(|c| {
                let mut r = c.response.lock().unwrap();
                if !r.headers_sent {
                    r.status_message = Some(message);
                }
            });
            Value::Null
        }),
    );

    // Status ───────────────────────────────────────────────────────────────
    vm.register_host_fn(
        "node:http",
        "set_status",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let code = args.first().map(|v| v.as_f64() as u16).unwrap_or(200);
            with_context(|c| {
                let mut r = c.response.lock().unwrap();
                if !r.headers_sent {
                    r.status = code;
                }
            });
            Value::Null
        }),
    );

    vm.register_host_fn(
        "node:http",
        "status",
        Box::new(|_ctx, _| {
            with_context(|c| Value::F64(c.response.lock().unwrap().status as f64))
                .unwrap_or(Value::F64(0.0))
        }),
    );

    // Headers ──────────────────────────────────────────────────────────────
    vm.register_host_fn(
        "node:http",
        "set_header",
        Box::new(|_ctx, args| {
            let name = string_arg(args, 0);
            let value = string_arg(args, 1);
            with_context(|c| {
                let mut r = c.response.lock().unwrap();
                if r.headers_sent {
                    return;
                }
                r.headers.retain(|(n, _)| !n.eq_ignore_ascii_case(&name));
                r.headers.push((name, value));
            });
            Value::Null
        }),
    );

    vm.register_host_fn(
        "node:http",
        "add_header",
        Box::new(|_ctx, args| {
            let name = string_arg(args, 0);
            let value = string_arg(args, 1);
            with_context(|c| {
                let mut r = c.response.lock().unwrap();
                if r.headers_sent {
                    return;
                }
                r.headers.push((name, value));
            });
            Value::Null
        }),
    );

    vm.register_host_fn(
        "node:http",
        "remove_header",
        Box::new(|_ctx, args| {
            let name = string_arg(args, 0);
            with_context(|c| {
                let mut r = c.response.lock().unwrap();
                if r.headers_sent {
                    return;
                }
                r.headers.retain(|(n, _)| !n.eq_ignore_ascii_case(&name));
            });
            Value::Null
        }),
    );

    vm.register_host_fn(
        "node:http",
        "has_header",
        Box::new(|_ctx, args| {
            let name = string_arg(args, 0);
            with_context(|c| {
                let r = c.response.lock().unwrap();
                Value::Bool(r.headers.iter().any(|(n, _)| n.eq_ignore_ascii_case(&name)))
            })
            .unwrap_or(Value::Bool(false))
        }),
    );

    vm.register_host_fn(
        "node:http",
        "headers_sent",
        Box::new(|_ctx, _| {
            with_context(|c| Value::Bool(c.response.lock().unwrap().headers_sent))
                .unwrap_or(Value::Bool(false))
        }),
    );

    // Body ─────────────────────────────────────────────────────────────────
    vm.register_host_fn(
        "node:http",
        "write",
        Box::new(|_ctx, args| {
            let bytes = match args.first() {
                Some(Value::String(s)) => s.as_bytes().to_vec(),
                Some(other) => format!("{}", other).into_bytes(),
                None => Vec::new(),
            };
            with_context(|c| {
                c.response.lock().unwrap().write_bytes(bytes);
            });
            Value::Null
        }),
    );

    vm.register_host_fn(
        "node:http",
        "end",
        Box::new(|_ctx, _| {
            with_context(|c| {
                c.response.lock().unwrap().end();
            });
            Value::Null
        }),
    );

    vm.register_host_fn(
        "node:http",
        "flush",
        Box::new(|_ctx, _| {
            // Phase 1: writes are unbuffered on the host side (each write_bytes
            // pushes a Data message). Explicit flush is a no-op for now.
            // Phase 2: add an ~8 KiB coalescing buffer and wire flush() to push
            // its contents.
            Value::Null
        }),
    );
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
fn _ensure_arc_in_scope(s: &str) -> Arc<str> {
    Arc::from(s)
}

/// Flatten `writeHead`'s optional headers argument.
///
/// Node accepts either an object (`{'Content-Type': 'text/plain'}`) or a flat
/// `[k, v, k, v]` array there, so both forms are handled.
fn header_pairs(value: Option<&Value>) -> Vec<(String, String)> {
    use vybe_runtime::value::ObjectKind;
    let Some(Value::Object(object)) = value else {
        return Vec::new();
    };
    let object = object.lock().unwrap();
    match &object.kind {
        ObjectKind::Array(items) => items
            .chunks(2)
            .filter(|pair| pair.len() == 2)
            .map(|pair| (as_text(&pair[0]), as_text(&pair[1])))
            .collect(),
        ObjectKind::Map(entries) => entries
            .iter()
            .map(|(k, v)| (as_text(k), as_text(v)))
            .collect(),
        _ => object
            .properties
            .iter()
            .map(|(k, v)| (k.to_string(), as_text(v)))
            .collect(),
    }
}

fn as_text(value: &Value) -> String {
    match value {
        Value::String(text) => text.to_string(),
        other => format!("{}", other),
    }
}
