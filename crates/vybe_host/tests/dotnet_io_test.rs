use std::fs;
use std::net::TcpListener;
use std::sync::Arc;
use vybe_bytecode::value::Value;
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn call_import(module: &str, name: &str, pre_stack: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<dotnet-io-test>");
    let import_idx = chunk.add_import(module, name);
    let argc = pre_stack.len() as u8;
    for value in pre_stack {
        let constant = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, constant, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn free_port() -> u16 {
    let listener = TcpListener::bind("127.0.0.1:0").expect("bind ephemeral port");
    listener.local_addr().expect("local addr").port()
}

#[test]
fn dotnet_io_file_reader_writer_roundtrip() {
    let path = std::env::temp_dir().join(format!("vybe-dotnet-io-{}.txt", std::process::id()));
    let path_value = Value::String(Arc::from(path.to_string_lossy().as_ref()));

    let writer = call_import("dotnet:io", "streamWriterNew", vec![path_value.clone()]);
    assert!(matches!(call_import("dotnet:io", "streamWriterWrite", vec![writer.clone(), Value::String(Arc::from("hello"))]), Value::Null));
    assert!(matches!(call_import("dotnet:io", "streamWriterWriteLine", vec![writer.clone(), Value::String(Arc::from(" world"))]), Value::Null));
    assert!(matches!(call_import("dotnet:io", "streamWriterClose", vec![writer]), Value::Null));

    let reader = call_import("dotnet:io", "streamReaderNew", vec![path_value]);
    let all = call_import("dotnet:io", "streamReaderReadToEnd", vec![reader]);
    assert_eq!(all.as_str(), "hello world");

    let _ = fs::remove_file(path);
}

/// .NET StreamReader/StreamWriter wrapping a TCP stream — exercises
/// the dotnet:io shim against sockets created via the spec
/// `wasi:sockets/*` primitives. The .NET socket adapter
/// (`emitter::dotnet::core::sockets_adapter`) compiles `New
/// TcpClient(...)` / `listener.AcceptTcpClient()` etc. through these
/// same primitives, so this test validates the shared back-end.
#[test]
fn dotnet_io_streams_wrap_tcp_client_streams() {
    let port = free_port();

    // Listener: create-socket → start-bind → finish-bind → start-listen → finish-listen
    let listener = call_import("wasi:sockets/tcp-create-socket", "create-tcp-socket",
        vec![Value::String(Arc::from("ipv4"))]);
    let bind_addr = Value::String(Arc::from(format!("0.0.0.0:{}", port).as_str()));
    let _ = call_import("wasi:sockets/tcp", "start-bind",
        vec![listener.clone(), Value::Null, bind_addr]);
    let _ = call_import("wasi:sockets/tcp", "finish-bind", vec![listener.clone()]);
    let _ = call_import("wasi:sockets/tcp", "start-listen", vec![listener.clone()]);
    let _ = call_import("wasi:sockets/tcp", "finish-listen", vec![listener.clone()]);

    // Client: create-socket + start-connect (synchronous in our impl)
    let client = call_import("wasi:sockets/tcp-create-socket", "create-tcp-socket",
        vec![Value::String(Arc::from("ipv4"))]);
    let connect_addr = Value::String(Arc::from(format!("127.0.0.1:{}", port).as_str()));
    let _ = call_import("wasi:sockets/tcp", "start-connect",
        vec![client.clone(), Value::Null, connect_addr]);

    // Accept (poll until pending in non-blocking mode). The wasi:sockets
    // accept returns [client_socket, input_stream, output_stream] | null.
    let accepted_array = {
        let mut last = Value::Null;
        for _ in 0..100 {
            last = call_import("wasi:sockets/tcp", "accept", vec![listener.clone()]);
            if !matches!(last, Value::Null) {
                break;
            }
            std::thread::sleep(std::time::Duration::from_millis(5));
        }
        last
    };
    let arr = match &accepted_array {
        Value::Object(o) => o.clone(),
        _ => panic!("accept returned non-array"),
    };
    let (accepted_socket, accepted_in_stream) = {
        let lo = arr.lock().unwrap();
        if let vybe_bytecode::value::ObjectKind::Array(ref elems) = lo.kind {
            (
                elems.first().cloned().unwrap_or(Value::Null),
                elems.get(1).cloned().unwrap_or(Value::Null),
            )
        } else {
            (Value::Null, Value::Null)
        }
    };

    // .NET StreamReader / StreamWriter wrap stream resources by `__type`
    // (`InputStream` / `OutputStream`) — exactly the shape WASI's
    // `tcp.accept` and `tcp.finish-connect` return. Pull the
    // output-stream half from `finish-connect(client)` and the
    // input-stream half from `accept`'s [1] element. The dotnet:io
    // shim retires next; this is the last test exercising it.
    let connect_streams = call_import("wasi:sockets/tcp", "finish-connect", vec![client.clone()]);
    let client_out_stream = match &connect_streams {
        Value::Object(o) => {
            let lo = o.lock().unwrap();
            if let vybe_bytecode::value::ObjectKind::Array(ref elems) = lo.kind {
                elems.get(1).cloned().unwrap_or(Value::Null)
            } else {
                Value::Null
            }
        }
        _ => Value::Null,
    };

    let writer = call_import("dotnet:io", "streamWriterNew", vec![client_out_stream]);
    let reader = call_import("dotnet:io", "streamReaderNew", vec![accepted_in_stream]);

    let _ = call_import("dotnet:io", "streamWriterWriteLine", vec![writer, Value::String(Arc::from("ping"))]);
    let line = call_import("dotnet:io", "streamReaderReadLine", vec![reader]);
    assert_eq!(line.as_str(), "ping");

    // Cleanup via wasi:sockets/tcp.shutdown
    let _ = call_import("wasi:sockets/tcp", "shutdown",
        vec![client, Value::String(Arc::from("both"))]);
    let _ = call_import("wasi:sockets/tcp", "shutdown",
        vec![accepted_socket, Value::String(Arc::from("both"))]);
    let _ = call_import("wasi:sockets/tcp", "shutdown",
        vec![listener, Value::String(Arc::from("both"))]);
}