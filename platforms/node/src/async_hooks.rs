//! `node:async_hooks` — Node.js async context tracking.
//!
//! Reference: <https://nodejs.org/api/async_hooks.html>.
//!
//! In Vybe's synchronous VM, async context tracking is best-effort:
//! AsyncLocalStorage state is stored on the object itself, and async IDs
//! are monotonically assigned counters (no true async propagation).

use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};
use vybe_runtime::VM;
use vybe_runtime::value::{Object, Value};

static NEXT_ASYNC_ID: AtomicU64 = AtomicU64::new(1);

fn next_async_id() -> u64 {
    NEXT_ASYNC_ID.fetch_add(1, Ordering::Relaxed)
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn get_prop(obj: &Value, key: &str) -> Value {
    if let Value::Object(o) = obj {
        o.lock()
            .unwrap()
            .properties
            .get(key)
            .cloned()
            .unwrap_or(Value::Undefined)
    } else {
        Value::Undefined
    }
}

fn set_prop(obj: &Value, key: &str, val: Value) {
    if let Value::Object(o) = obj {
        o.lock().unwrap().properties.insert(key.to_string(), val);
    }
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "node:async_hooks",
        "executionAsyncId",
        Box::new(|_ctx, _args| Value::I32(1)),
    );

    vm.register_host_fn(
        "node:async_hooks",
        "triggerAsyncId",
        Box::new(|_ctx, _args| Value::I32(0)),
    );

    vm.register_host_fn(
        "node:async_hooks",
        "executionAsyncResource",
        Box::new(|_ctx, _args| Value::Object(vybe_runtime::heap::alloc(Object::new()))),
    );

    // ── AsyncLocalStorage ─────────────────────────────────────────────────────

    vm.register_host_fn(
        "node:async_hooks",
        "AsyncLocalStorage",
        Box::new(|_ctx, _args| {
            let mut o = Object::new();
            o.properties.insert("__store".into(), Value::Undefined);
            o.properties.insert("__disabled".into(), Value::Bool(false));
            Value::Object(vybe_runtime::heap::alloc(o))
        }),
    );

    vm.register_host_fn(
        "node:async_hooks",
        "alsGetStore",
        Box::new(|_ctx, args| {
            let als = args.first().cloned().unwrap_or(Value::Undefined);
            get_prop(&als, "__store")
        }),
    );

    vm.register_host_fn(
        "node:async_hooks",
        "alsRun",
        Box::new(|_ctx, args| {
            let als = args.first().cloned().unwrap_or(Value::Undefined);
            let store_val = args.get(1).cloned().unwrap_or(Value::Undefined);
            // Temporarily set store (callback not invocable from host — return null)
            set_prop(&als, "__store", store_val);
            Value::Null
        }),
    );

    vm.register_host_fn(
        "node:async_hooks",
        "alsEnterWith",
        Box::new(|_ctx, args| {
            let als = args.first().cloned().unwrap_or(Value::Undefined);
            let val = args.get(1).cloned().unwrap_or(Value::Undefined);
            set_prop(&als, "__store", val);
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        "node:async_hooks",
        "alsExit",
        Box::new(|_ctx, _args| Value::Undefined),
    );

    vm.register_host_fn(
        "node:async_hooks",
        "alsDisable",
        Box::new(|_ctx, args| {
            let als = args.first().cloned().unwrap_or(Value::Undefined);
            set_prop(&als, "__disabled", Value::Bool(true));
            Value::Undefined
        }),
    );

    // ── AsyncResource ─────────────────────────────────────────────────────────

    vm.register_host_fn(
        "node:async_hooks",
        "AsyncResource",
        Box::new(|_ctx, args| {
            let type_name = match args.first() {
                Some(Value::String(s)) => s.to_string(),
                _ => String::new(),
            };
            let async_id = next_async_id();
            let mut o = Object::new();
            o.properties.insert("__type".into(), s(&type_name));
            o.properties
                .insert("__asyncId".into(), Value::I64(async_id as i64));
            o.properties
                .insert("__triggerAsyncId".into(), Value::I32(0));
            Value::Object(vybe_runtime::heap::alloc(o))
        }),
    );

    vm.register_host_fn(
        "node:async_hooks",
        "asyncResourceAsyncId",
        Box::new(|_ctx, args| {
            let res = args.first().cloned().unwrap_or(Value::Undefined);
            match get_prop(&res, "__asyncId") {
                Value::Undefined => Value::I32(1),
                v => v,
            }
        }),
    );

    vm.register_host_fn(
        "node:async_hooks",
        "asyncResourceTriggerAsyncId",
        Box::new(|_ctx, args| {
            let res = args.first().cloned().unwrap_or(Value::Undefined);
            match get_prop(&res, "__triggerAsyncId") {
                Value::Undefined => Value::I32(0),
                v => v,
            }
        }),
    );

    vm.register_host_fn(
        "node:async_hooks",
        "asyncResourceRun",
        Box::new(|_ctx, _args| Value::Null),
    );

    vm.register_host_fn(
        "node:async_hooks",
        "asyncResourceBind",
        Box::new(|_ctx, _args| Value::Null),
    );

    // ── createHook / hookEnable / hookDisable ────────────────────────────────

    vm.register_host_fn(
        "node:async_hooks",
        "createHook",
        Box::new(|_ctx, _args| {
            let mut o = Object::new();
            o.properties.insert("__enabled".into(), Value::Bool(false));
            Value::Object(vybe_runtime::heap::alloc(o))
        }),
    );

    vm.register_host_fn(
        "node:async_hooks",
        "hookEnable",
        Box::new(|_ctx, args| {
            let hook = args.first().cloned().unwrap_or(Value::Undefined);
            set_prop(&hook, "__enabled", Value::Bool(true));
            match &hook {
                Value::Object(_) => hook,
                _ => Value::Object(vybe_runtime::heap::alloc(Object::new())),
            }
        }),
    );

    vm.register_host_fn(
        "node:async_hooks",
        "hookDisable",
        Box::new(|_ctx, args| {
            let hook = args.first().cloned().unwrap_or(Value::Undefined);
            set_prop(&hook, "__enabled", Value::Bool(false));
            match &hook {
                Value::Object(_) => hook,
                _ => Value::Object(vybe_runtime::heap::alloc(Object::new())),
            }
        }),
    );

    // ── asyncWrapProviders ────────────────────────────────────────────────────

    vm.register_host_fn(
        "node:async_hooks",
        "asyncWrapProviders",
        Box::new(|_ctx, _args| {
            let mut o = Object::new();
            // Numeric constants matching Node's async provider IDs
            for (name, id) in [
                ("NONE", 0i32),
                ("CRYPTO", 1),
                ("DNSCHANNEL", 2),
                ("ELDHISTOGRAM", 3),
                ("FILEHANDLE", 4),
                ("FILEHANDLECLOSEREQ", 5),
                ("FIXEDSIZEBLOBCOPY", 6),
                ("FSEVENTWRAP", 7),
                ("FSREQCALLBACK", 8),
                ("FSREQPROMISE", 9),
                ("GETADDRINFOREQWRAP", 10),
                ("GETNAMEINFOREQWRAP", 11),
                ("HEAPSNAPSHOT", 12),
                ("HTTP2SESSION", 13),
                ("HTTP2STREAM", 14),
                ("HTTP2PING", 15),
                ("HTTP2SETTINGS", 16),
                ("HTTPINCOMINGMESSAGE", 17),
                ("HTTPCLIENTREQUEST", 18),
                ("JSSTREAM", 19),
                ("JSUDPWRAP", 20),
                ("MESSAGEPORT", 21),
                ("PIPECONNECTWRAP", 22),
                ("PIPESERVERWRAP", 23),
                ("PIPEWRAP", 24),
                ("PROCESSWRAP", 25),
                ("PROMISE", 26),
                ("QUERYWRAP", 27),
                ("SHUTDOWNWRAP", 28),
                ("SIGNALWRAP", 29),
                ("STATWATCHER", 30),
                ("STREAMPIPE", 31),
                ("TCPCONNECTWRAP", 32),
                ("TCPSERVERWRAP", 33),
                ("TCPWRAP", 34),
                ("TTYWRAP", 35),
                ("UDPSENDWRAP", 36),
                ("UDPWRAP", 37),
                ("SIGINTWATCHDOG", 38),
                ("WORKER", 39),
                ("WORKERHEAPSNAPSHOT", 40),
                ("WRITEWRAP", 41),
                ("ZLIB", 42),
            ] {
                o.properties.insert(name.into(), Value::I32(id));
            }
            Value::Object(vybe_runtime::heap::alloc(o))
        }),
    );
}
