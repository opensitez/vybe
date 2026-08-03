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
    // `exit(status: result<(), ()>)` carries no number: ok is success, err is
    // failure. Anything truthy passed here is the error arm, hence 1.
    vm.register_host_fn(
        "wasi:cli/exit",
        "exit",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let failed = args.first().is_some_and(|v| v.as_bool());
            ctx.request_exit_with_code(if failed { 1 } else { 0 });
            Value::Null
        }),
    );

    // `exit-with-code(status: u8)` is the arm that carries a real status. The
    // code was being dropped here (`_args`), which is why `sys.exit(3)`,
    // `System.exit(4)` and `halt(2)` all produced 0. The VM only *carries* it;
    // turning it into a process status stays with the embedder.
    vm.register_host_fn(
        "wasi:cli/exit",
        "exit-with-code",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let code = match args.first() {
                Some(Value::I32(_) | Value::I64(_) | Value::F32(_) | Value::F64(_) | Value::Bool(_)) => {
                    args[0].as_i32()
                }
                // Python's `sys.exit("message")` prints the object to stderr and
                // exits 1; a bare `sys.exit()` / `sys.exit(None)` exits 0.
                Some(Value::Null | Value::Undefined) | None => 0,
                Some(other) => {
                    eprintln!("{}", other);
                    1
                }
            };
            ctx.request_exit_with_code(code);
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

    // wasi:cli/stdin — 0.3 `read-via-stream: func() -> tuple<stream<u8>,
    // future<result<_, error-code>>>` (`proposals/cli/wit/stdio.wit`).
    //
    // The stream carries stdin's bytes and the future signals how the read
    // ended: resolved with success at clean EOF, with an `error-code` if the
    // read failed. When stdin is a terminal the stream is closed empty rather
    // than blocking the process waiting for a key — interactive reads are the
    // 0.2 `get-stdin` + `input-stream.blocking-read` path, which stays bound
    // above and is what a prompting program uses.
    vm.register_host_fn(
        "wasi:cli/stdin",
        "read-via-stream",
        Box::new(|ctx: &mut HostContext, _args: &[Value]| {
            use std::io::{IsTerminal, Read};

            let (stream_val, stream_id) = ctx.create_stream();
            let mut failure: Option<&str> = None;
            if !std::io::stdin().is_terminal() {
                let mut buffer = Vec::new();
                match std::io::stdin().read_to_end(&mut buffer) {
                    Ok(_) => {
                        for byte in &buffer {
                            ctx.stream_push(stream_id, Value::I32(*byte as i32));
                        }
                    }
                    // `error-code` in `wasi:cli/types`: io, illegal-byte-sequence, pipe.
                    Err(error) => {
                        failure = Some(match error.kind() {
                            std::io::ErrorKind::BrokenPipe => "pipe",
                            std::io::ErrorKind::InvalidData => "illegal-byte-sequence",
                            _ => "io",
                        })
                    }
                }
            }
            ctx.stream_close(stream_id);

            let (future_val, future_id) = ctx.create_future();
            let outcome = match failure {
                Some(code) => Value::String(Arc::from(code)),
                None => Value::Null,
            };
            ctx.resolve_future(future_id, outcome);
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
                stream_val, future_val,
            ])))
        }),
    );

    // wasi:cli/terminal-{stdin,stdout,stderr} — 0.3
    // `get-terminal-*: func() -> option<terminal-{input,output}>`
    // (`proposals/cli/wit/terminal.wit`). The resources carry no methods yet;
    // their presence IS the answer to "is this stream a terminal".
    for (module, name, kind, fd) in [
        ("wasi:cli/terminal-stdin", "get-terminal-stdin", "terminal-input", 0),
        (
            "wasi:cli/terminal-stdout",
            "get-terminal-stdout",
            "terminal-output",
            1,
        ),
        (
            "wasi:cli/terminal-stderr",
            "get-terminal-stderr",
            "terminal-output",
            2,
        ),
    ] {
        vm.register_host_fn(
            module,
            name,
            Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
                use std::io::IsTerminal;
                let attached = match fd {
                    0 => std::io::stdin().is_terminal(),
                    1 => std::io::stdout().is_terminal(),
                    _ => std::io::stderr().is_terminal(),
                };
                if !attached {
                    return Value::Null;
                }
                let mut handle = Object::new();
                handle.properties.insert("fd".into(), Value::I32(fd));
                handle
                    .properties
                    .insert("__type".into(), Value::String(Arc::from(kind)));
                Value::Object(vybe_runtime::heap::alloc(handle))
            }),
        );
    }

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
