use std::sync::Arc;
use vybe_runtime::value::Object;
use vybe_runtime::{HostContext, VM, Value};

pub fn register(vm: &mut VM) {
    // wasi:logging/logging.log — WASI logging proposal
    // Signature: log(level: level, context: string, message: string)
    // Flexible arity: 1 arg = (info, "", msg); 2 args = (level, msg);
    // 3 args = (level, context, msg); N>3 args = (info, "", joined).
    // info/debug/trace → stdout; warn/error/critical → stderr.
    vm.register_host_fn(
        "wasi:logging/logging",
        "log",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let (level, message) = match args.len() {
                0 => ("info".to_string(), String::new()),
                1 => ("info".to_string(), format!("{}", args[0])),
                2 => (format!("{}", args[0]), format!("{}", args[1])),
                3 => (format!("{}", args[0]), format!("{}", args[2])),
                _ => {
                    let parts: Vec<String> = args.iter().map(|v| format!("{}", v)).collect();
                    ("info".to_string(), parts.join(" "))
                }
            };
            match level.as_str() {
                "warn" | "error" | "critical" => eprintln!("{}", message),
                _ => println!("{}", message),
            }
            Value::Null
        }),
    );

    // wasi:cli/exit — WASI 0.2.12+ spec interface. Per the component model,
    // `exit` terminates the GUEST instance and returns control to the embedder;
    // it must NOT tear down the host process. Signalling the VM to end the run
    // (request_exit) is the conformant behaviour — `std::process::exit` here
    // would kill the whole embedder (e.g. the test binary) on the first call.
    vm.register_host_fn(
        "wasi:cli/exit",
        "exit",
        Box::new(|ctx: &mut HostContext, _args: &[Value]| {
            ctx.request_exit();
            Value::Null
        }),
    );

    vm.register_host_fn(
        "wasi:cli/exit",
        "exit-with-code",
        Box::new(|ctx: &mut HostContext, _args: &[Value]| {
            ctx.request_exit();
            Value::Null
        }),
    );

    // wasi:cli/stdout|stderr|stdin — return a stream handle with an fd tag.
    // These are used by [method]output-stream.blocking-write-and-flush in wasi:io/streams.
    vm.register_host_fn(
        "wasi:cli/stdout",
        "get-stdout",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            let mut h = Object::new();
            h.properties.insert("fd".into(), Value::I32(1));
            h.properties
                .insert("__type".into(), Value::String(Arc::from("output-stream")));
            Value::Object(vybe_runtime::heap::alloc(h))
        }),
    );

    vm.register_host_fn(
        "wasi:cli/stderr",
        "get-stderr",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            let mut h = Object::new();
            h.properties.insert("fd".into(), Value::I32(2));
            h.properties
                .insert("__type".into(), Value::String(Arc::from("output-stream")));
            Value::Object(vybe_runtime::heap::alloc(h))
        }),
    );

    vm.register_host_fn(
        "wasi:cli/stdin",
        "get-stdin",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            let mut h = Object::new();
            h.properties.insert("fd".into(), Value::I32(0));
            h.properties
                .insert("__type".into(), Value::String(Arc::from("input-stream")));
            Value::Object(vybe_runtime::heap::alloc(h))
        }),
    );

    // wasi:cli/stdout|stderr — 0.3 write-via-stream
    vm.register_host_fn(
        "wasi:cli/stdout",
        "write-via-stream",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let stream_val = args.first().cloned().unwrap_or(Value::Null);
            let bytes = ctx.stream_drain(&stream_val);
            if !bytes.is_empty() {
                use std::io::Write;
                let _ = std::io::stdout().write_all(&bytes);
            }
            let (fut, fut_id) = ctx.create_future();
            ctx.resolve_future(fut_id, Value::Null);
            fut
        }),
    );

    vm.register_host_fn(
        "wasi:cli/stderr",
        "write-via-stream",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let stream_val = args.first().cloned().unwrap_or(Value::Null);
            let bytes = ctx.stream_drain(&stream_val);
            if !bytes.is_empty() {
                use std::io::Write;
                let _ = std::io::stderr().write_all(&bytes);
            }
            let (fut, fut_id) = ctx.create_future();
            ctx.resolve_future(fut_id, Value::Null);
            fut
        }),
    );
}
