//! `vybe:http/response.*` — per-request response write.
//!
//! All mutators silently no-op when no request context is installed.
//! Setting headers / status after the first body write is also a no-op
//! (matching PHP's `headers_sent` warning behavior).

use std::sync::Arc;
use vybe_bytecode::{VM, Value, HostContext};

use super::context::with_context;

pub fn register(vm: &mut VM) {
    // Status ───────────────────────────────────────────────────────────────
    vm.register_host_fn("vybe:http/response", "set_status", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let code = args.first().map(|v| v.as_f64() as u16).unwrap_or(200);
        with_context(|c| {
            let mut r = c.response.lock().unwrap();
            if !r.headers_sent { r.status = code; }
        });
        Value::Null
    }));

    vm.register_host_fn("vybe:http/response", "status", Box::new(|_ctx, _| {
        with_context(|c| Value::F64(c.response.lock().unwrap().status as f64))
            .unwrap_or(Value::F64(0.0))
    }));

    // Headers ──────────────────────────────────────────────────────────────
    vm.register_host_fn("vybe:http/response", "set_header", Box::new(|_ctx, args| {
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

    vm.register_host_fn("vybe:http/response", "add_header", Box::new(|_ctx, args| {
        let name = string_arg(args, 0);
        let value = string_arg(args, 1);
        with_context(|c| {
            let mut r = c.response.lock().unwrap();
            if r.headers_sent { return; }
            r.headers.push((name, value));
        });
        Value::Null
    }));

    vm.register_host_fn("vybe:http/response", "remove_header", Box::new(|_ctx, args| {
        let name = string_arg(args, 0);
        with_context(|c| {
            let mut r = c.response.lock().unwrap();
            if r.headers_sent { return; }
            r.headers.retain(|(n, _)| !n.eq_ignore_ascii_case(&name));
        });
        Value::Null
    }));

    vm.register_host_fn("vybe:http/response", "has_header", Box::new(|_ctx, args| {
        let name = string_arg(args, 0);
        with_context(|c| {
            let r = c.response.lock().unwrap();
            Value::Bool(r.headers.iter().any(|(n, _)| n.eq_ignore_ascii_case(&name)))
        }).unwrap_or(Value::Bool(false))
    }));

    vm.register_host_fn("vybe:http/response", "headers_sent", Box::new(|_ctx, _| {
        with_context(|c| Value::Bool(c.response.lock().unwrap().headers_sent))
            .unwrap_or(Value::Bool(false))
    }));

    // Body ─────────────────────────────────────────────────────────────────
    vm.register_host_fn("vybe:http/response", "write", Box::new(|_ctx, args| {
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

    vm.register_host_fn("vybe:http/response", "write_text", Box::new(|_ctx, args| {
        let text = string_arg(args, 0);
        with_context(|c| {
            c.response.lock().unwrap().write_bytes(text.into_bytes());
        });
        Value::Null
    }));

    vm.register_host_fn("vybe:http/response", "end", Box::new(|_ctx, _| {
        with_context(|c| {
            c.response.lock().unwrap().end();
        });
        Value::Null
    }));

    vm.register_host_fn("vybe:http/response", "flush", Box::new(|_ctx, _| {
        // Phase 1: writes are unbuffered on the host side (each write_bytes
        // pushes a Data message). Explicit flush is a no-op for now.
        // Phase 2: add an ~8 KiB coalescing buffer and wire flush() to push
        // its contents.
        Value::Null
    }));
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
