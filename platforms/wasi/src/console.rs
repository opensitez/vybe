use std::sync::Arc;
use vybe_runtime::value::Object;
use vybe_runtime::{HostContext, VM, Value};

fn read_chunks(
    mut input: impl std::io::Read,
    mut push: impl FnMut(Vec<u8>),
) -> std::io::Result<()> {
    let mut first = [0u8; 8192];
    for _ in 0..2 {
        loop {
            match input.read(&mut first) {
                Ok(0) => return Ok(()),
                Ok(n) => {
                    push(first[..n].to_vec());
                    break;
                }
                Err(error) if error.kind() == std::io::ErrorKind::Interrupted => continue,
                Err(error) => return Err(error),
            }
        }
    }
    let mut buffer = vec![0u8; 65536];
    loop {
        match input.read(&mut buffer) {
            Ok(0) => return Ok(()),
            Ok(n) => {
                buffer.truncate(n);
                push(buffer);
                buffer = vec![0u8; 65536];
            }
            Err(error) if error.kind() == std::io::ErrorKind::Interrupted => continue,
            Err(error) => return Err(error),
        }
    }
}

fn write_bytes(mut output: impl std::io::Write, bytes: &[u8]) -> Result<(), &'static str> {
    output.write_all(bytes).map_err(output_error)
}

fn output_error(error: std::io::Error) -> &'static str {
    match error.kind() {
        std::io::ErrorKind::BrokenPipe => "pipe",
        std::io::ErrorKind::InvalidData => "illegal-byte-sequence",
        _ => "io",
    }
}

fn flush_output(mut output: impl std::io::Write) -> Result<(), &'static str> {
    output.flush().map_err(output_error)
}

#[cfg(test)]
mod read_tests {
    use super::read_chunks;
    use std::io::{self, Cursor, Read};

    #[test]
    fn pipe_bytes_are_chunked_without_changing_the_payload() {
        let input: Vec<u8> = (0..131073).map(|i| (i % 251) as u8).collect();
        let mut chunks = Vec::new();
        read_chunks(Cursor::new(&input), |bytes| chunks.push(bytes)).unwrap();
        assert!(chunks.len() >= 3);
        assert!(chunks.iter().all(|chunk| chunk.len() <= 65536));
        assert_eq!(chunks.concat(), input);
    }

    struct InterruptedThenRead<R>(bool, R);

    impl<R: Read> Read for InterruptedThenRead<R> {
        fn read(&mut self, bytes: &mut [u8]) -> io::Result<usize> {
            if !self.0 {
                self.0 = true;
                return Err(io::ErrorKind::Interrupted.into());
            }
            self.1.read(bytes)
        }
    }

    #[test]
    fn interrupted_read_retries() {
        let mut output = Vec::new();
        read_chunks(InterruptedThenRead(false, Cursor::new(b"hello")), |bytes| {
            output.extend_from_slice(&bytes)
        })
        .unwrap();
        assert_eq!(output, b"hello");
    }

    #[test]
    fn empty_and_short_input_do_not_need_bulk_chunks() {
        let mut chunks = Vec::new();
        read_chunks(Cursor::new([]), |bytes| chunks.push(bytes)).unwrap();
        assert!(chunks.is_empty());

        read_chunks(Cursor::new(b"hello"), |bytes| chunks.push(bytes)).unwrap();
        assert_eq!(chunks, vec![b"hello".to_vec()]);
    }

    #[test]
    fn error_after_initial_bytes_keeps_them() {
        struct FirstThenError(bool);

        impl Read for FirstThenError {
            fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
                if std::mem::replace(&mut self.0, true) {
                    Err(io::ErrorKind::PermissionDenied.into())
                } else {
                    output[..3].copy_from_slice(b"abc");
                    Ok(3)
                }
            }
        }

        let mut chunks = Vec::new();
        let error = read_chunks(FirstThenError(false), |bytes| chunks.push(bytes)).unwrap_err();
        assert_eq!(error.kind(), io::ErrorKind::PermissionDenied);
        assert_eq!(chunks, vec![b"abc".to_vec()]);
    }

    #[test]
    #[ignore = "local small-stdin allocation benchmark"]
    fn bench_small_input_against_bulk_buffer() {
        use std::time::Instant;

        for input in [b"".as_slice(), b"hello".as_slice()] {
            for _ in 0..3 {
                let start = Instant::now();
                for _ in 0..100_000 {
                    let mut chunks = Vec::new();
                    read_chunks(Cursor::new(input), |bytes| chunks.push(bytes)).unwrap();
                    std::hint::black_box(chunks);
                }
                let small = start.elapsed();

                let start = Instant::now();
                for _ in 0..100_000 {
                    let mut source = Cursor::new(input);
                    let mut chunks = Vec::new();
                    loop {
                        let mut buffer = vec![0u8; 65536];
                        let count = source.read(&mut buffer).unwrap();
                        if count == 0 {
                            break;
                        }
                        buffer.truncate(count);
                        chunks.push(buffer);
                    }
                    std::hint::black_box(chunks);
                }
                eprintln!("stdin {} bytes: small {small:?}, bulk {:?}", input.len(), start.elapsed());
            }
        }
    }

    #[test]
    #[ignore = "local stdin bulk-read allocation benchmark"]
    fn bench_owned_chunks_against_copied_chunks() {
        use std::time::Instant;

        let input = vec![123u8; 16 * 1024 * 1024];
        for _ in 0..3 {
            let mut total = 0;
            let start = Instant::now();
            read_chunks(Cursor::new(&input), |bytes| {
                total += bytes.len();
                std::hint::black_box(bytes);
            })
            .unwrap();
            let owned = start.elapsed();
            assert_eq!(total, input.len());

            let mut source = Cursor::new(&input);
            let mut buffer = [0u8; 65536];
            let start = Instant::now();
            loop {
                let read = source.read(&mut buffer).unwrap();
                if read == 0 {
                    break;
                }
                std::hint::black_box(buffer[..read].to_vec());
            }
            eprintln!("stdin 16 MiB owned {owned:?}, copied {:?}", start.elapsed());
        }
    }
}

pub fn register(vm: &mut VM) {
    // wasi:logging/logging.log — the WASI logging proposal, STRICT:
    // `log(level: level, context: string, message: string)`. Exactly one
    // signature, no arity heuristics — the old sniffing treated 2- and
    // 3-arg CONSOLE calls as (level, msg)/(level, ctx, msg) and silently
    // dropped arguments (`console.log("a","b")` printed only `b`).
    // Console/print-shaped variadic output lives at `web:console.log`
    // (WHATWG — variadic by spec). info/debug/trace → stdout;
    // warn/error/critical → stderr.
    vm.register_host_fn(
        "wasi:logging/logging",
        "log",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let message = args.get(2);
            let stderr = matches!(args.first(), Some(Value::String(level))
                if matches!(level.as_ref(), "warn" | "error" | "critical"));
            match (stderr, message) {
                (true, Some(message)) => eprintln!("{message}"),
                (false, Some(message)) => println!("{message}"),
                (true, None) => eprintln!(),
                (false, None) => println!(),
            }
            Value::Null
        }),
    );

    // wasi:cli/exit — `wasi:cli@0.3.1`. Per the component model,
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
                Some(
                    Value::I32(_) | Value::I64(_) | Value::F32(_) | Value::F64(_) | Value::Bool(_),
                ) => args[0].as_i32(),
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

    // `get-stdout`, `get-stderr` and `get-stdin` USED TO BE REGISTERED HERE.
    //
    // Each answered a stream HANDLE tagged with an fd, for the 0.2 pair

    // wasi:cli/stdin — 0.3 `read-via-stream: func() -> tuple<stream<u8>,
    // future<result<_, error-code>>>` (`proposals/cli/wit/stdio.wit`).
    //
    // The stream carries stdin's bytes and the future signals how the read
    // ended: resolved with success at clean EOF, with an `error-code` if the
    // read failed.
    //
    // How MUCH arrives per call depends on what stdin is, and both answers are
    // honest ones:
    //
    //   - A pipe or a file is read to EOF in bounded chunks; the next call
    //     then sees EOF.
    //   - A TERMINAL has nothing until somebody types, so this blocks for one
    //     line and closes. It used to close EMPTY here on the grounds that
    //     interactive reads were the 0.2 `get-stdin` + `blocking-read` path —
    vm.register_host_fn(
        "wasi:cli/stdin",
        "read-via-stream",
        Box::new(|ctx: &mut HostContext, _args: &[Value]| {
            use std::io::{BufRead, IsTerminal};

            let (stream_val, stream_id) = ctx.create_stream();
            let mut failure: Option<&str> = None;
            if std::io::stdin().is_terminal() {
                let mut buffer = Vec::new();
                match std::io::stdin().lock().read_until(b'\n', &mut buffer) {
                    Ok(_) => {
                        ctx.stream_push_bytes(stream_id, buffer);
                    }
                    Err(error) => {
                        failure = Some(match error.kind() {
                            std::io::ErrorKind::BrokenPipe => "pipe",
                            std::io::ErrorKind::InvalidData => "illegal-byte-sequence",
                            _ => "io",
                        })
                    }
                }
            } else {
                match read_chunks(std::io::stdin().lock(), |bytes| {
                    ctx.stream_push_bytes(stream_id, bytes);
                }) {
                    Ok(()) => {}
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
                Some(code) => super::filesystem::err(code),
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
        (
            "wasi:cli/terminal-stdin",
            "get-terminal-stdin",
            "terminal-input",
            0,
        ),
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
            let (fut, fut_id) = ctx.create_future();
            let result = args.first().ok_or("invalid stream").and_then(|stream| {
                ctx.consume_byte_stream(stream, fut_id, Box::new(|bytes| {
                    if let Some(bytes) = bytes {
                        write_bytes(std::io::stdout().lock(), bytes)
                            .map_err(super::filesystem::err)?;
                        vybe_runtime::debugger::record_stdout_write();
                    } else {
                        flush_output(std::io::stdout().lock())
                            .map_err(super::filesystem::err)?;
                    }
                    Ok(())
                }))
            });
            if result.is_err() { ctx.resolve_future(fut_id, super::filesystem::err("io")); }
            fut
        }),
    );

    vm.register_host_fn(
        "wasi:cli/stderr",
        "write-via-stream",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let (fut, fut_id) = ctx.create_future();
            let result = args.first().ok_or("invalid stream").and_then(|stream| {
                ctx.consume_byte_stream(stream, fut_id, Box::new(|bytes| {
                    if let Some(bytes) = bytes {
                        write_bytes(std::io::stderr().lock(), bytes)
                            .map_err(super::filesystem::err)?;
                    } else {
                        flush_output(std::io::stderr().lock())
                            .map_err(super::filesystem::err)?;
                    }
                    Ok(())
                }))
            });
            if result.is_err() { ctx.resolve_future(fut_id, super::filesystem::err("io")); }
            fut
        }),
    );
}

#[cfg(test)]
mod stream_sink_tests {
    use super::*;
    use vybe_runtime::{Chunk, Op};
    use vybe_runtime::event_loop::FuturePhase;

    #[test]
    fn invalid_stdio_stream_resolves_canonical_error_arm() {
        for module in ["wasi:cli/stdout", "wasi:cli/stderr"] {
            let mut vm = VM::new();
            register(&mut vm);
            let mut chunk = Chunk::new("invalid-stdio-stream");
            let sink = chunk.add_import(module, "write-via-stream");
            chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
            chunk.emit_call(sink, 1, 0);
            chunk.emit_op(Op::RETURN, 0);
            assert!(matches!(vm.run(vec![chunk]).unwrap(), Value::I32(_)));
            let events = vm.event_loop.borrow();
            let completion = events.future_states.values().next().unwrap();
            assert_eq!(completion.phase, FuturePhase::Resolved);
            let value = completion.value.as_ref().unwrap();
            let (case, payload) = vybe_runtime::canon_value::result_parts_public(value);
            assert_eq!(case, 1);
            assert!(matches!(payload, Value::String(ref code) if code.as_ref() == "io"));
            let ty = vybe_runtime::component::ValType::Result(None, Some(Box::new(
                vybe_runtime::component::ValType::Variant(vec![("io".into(), None)]),
            )));
            let memory = vybe_runtime::shared_memory::SharedMemory::new(8);
            vybe_runtime::canon_value::store(&memory, value, &ty, 0).unwrap();
            assert_eq!(memory.read_array::<2>(0).unwrap(), [1, 0]);
        }
    }

    #[test]
    fn real_stdio_sink_stays_pending_until_the_guest_writer_closes() {
        for module in ["wasi:cli/stdout", "wasi:cli/stderr"] {
            let mut vm = VM::new();
            register(&mut vm);
            let mut start = Chunk::new("stdio-start-open-stream");
            let new = start.add_import("canon", "stream.new");
            let sink = start.add_import(module, "write-via-stream");
            let writer = start.alloc_scratch(1);
            start.local_count = start.local_count.max(start.scratch_high_water);
            let packed = start.i64_scratch();
            start.emit_call(new, 0, 0);
            start.emit_op_u16(Op::LOCAL_SET, packed, 0);
            start.emit_op_u16(Op::LOCAL_GET, packed, 0);
            start.emit_i64_const(32, 0);
            start.emit_op(Op::I64_SHR_U, 0);
            start.emit_op(Op::I32_WRAP_I64, 0);
            start.emit_op_u16(Op::LOCAL_SET, writer, 0);
            start.emit_op_u16(Op::LOCAL_GET, packed, 0);
            start.emit_op(Op::I32_WRAP_I64, 0);
            start.emit_call(sink, 1, 0);
            start.emit_op(Op::DROP, 0);
            start.emit_op_u16(Op::LOCAL_GET, writer, 0);
            start.emit_op(Op::RETURN, 0);
            let writer = vm.run(vec![start]).unwrap().as_i32();
            {
                let events = vm.event_loop.borrow();
                assert_eq!(events.future_states.len(), 1);
                assert_eq!(events.future_states.values().next().unwrap().phase, FuturePhase::Pending);
            }
            let mut finish = Chunk::new("stdio-close-stream");
            let close = finish.add_import("canon", "stream.drop-writable");
            finish.emit_i32_const(writer, 0);
            finish.emit_call(close, 1, 0);
            finish.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
            finish.emit_op(Op::RETURN, 0);
            vm.run(vec![finish]).unwrap();
            let events = vm.event_loop.borrow();
            assert_eq!(events.future_states.len(), 1);
            let completion = events.future_states.values().next().unwrap();
            assert_eq!(completion.phase, FuturePhase::Resolved);
            assert!(matches!(completion.value, Some(Value::Null)));
        }
    }
}

#[cfg(test)]
mod write_tests {
    use super::{flush_output, write_bytes};
    use std::io::{self, Write};

    struct BrokenWriter;

    impl Write for BrokenWriter {
        fn write(&mut self, _: &[u8]) -> io::Result<usize> {
            Err(io::Error::from(io::ErrorKind::BrokenPipe))
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[test]
    fn output_errors_keep_the_wasi_error_code() {
        assert_eq!(write_bytes(BrokenWriter, b"data"), Err("pipe"));
        assert_eq!(write_bytes(io::sink(), b"data"), Ok(()));
    }

    #[test]
    fn eof_flush_preserves_errors_and_reaches_the_writer() {
        struct FlushWriter(bool);
        impl io::Write for FlushWriter {
            fn write(&mut self, bytes: &[u8]) -> io::Result<usize> { Ok(bytes.len()) }
            fn flush(&mut self) -> io::Result<()> {
                self.0 = true;
                Ok(())
            }
        }
        let mut writer = FlushWriter(false);
        assert_eq!(flush_output(&mut writer), Ok(()));
        assert!(writer.0);
        struct FlushFailure;
        impl io::Write for FlushFailure {
            fn write(&mut self, bytes: &[u8]) -> io::Result<usize> { Ok(bytes.len()) }
            fn flush(&mut self) -> io::Result<()> {
                Err(io::Error::from(io::ErrorKind::BrokenPipe))
            }
        }
        assert_eq!(flush_output(FlushFailure), Err("pipe"));
    }
}
