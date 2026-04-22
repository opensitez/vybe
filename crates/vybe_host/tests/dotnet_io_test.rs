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

#[test]
fn dotnet_io_streams_wrap_tcp_client_streams() {
    let port = free_port();

    let listener = call_import("vybe:net", "tcpListenerNew", vec![Value::I32(port as i32)]);
    let _ = call_import("vybe:net", "tcpListenerStart", vec![listener.clone()]);
    let client = call_import("vybe:net", "tcpConnect", vec![Value::String(Arc::from("127.0.0.1")), Value::I32(port as i32)]);
    let accepted = {
        let mut last = Value::Null;
        for _ in 0..100 {
            last = call_import("vybe:net", "tcpListenerAccept", vec![listener.clone()]);
            if !matches!(last, Value::Null) {
                break;
            }
            std::thread::sleep(std::time::Duration::from_millis(5));
        }
        last
    };

    let client_stream = call_import("vybe:net", "tcpGetStream", vec![client.clone()]);
    let writer = call_import("dotnet:io", "streamWriterNew", vec![client_stream]);
    let reader = call_import("dotnet:io", "streamReaderNew", vec![accepted.clone()]);

    let _ = call_import("dotnet:io", "streamWriterWriteLine", vec![writer, Value::String(Arc::from("ping"))]);
    let line = call_import("dotnet:io", "streamReaderReadLine", vec![reader]);
    assert_eq!(line.as_str(), "ping");

    let _ = call_import("vybe:net", "tcpClose", vec![client]);
    let _ = call_import("vybe:net", "tcpClose", vec![accepted]);
    let _ = call_import("vybe:net", "tcpListenerStop", vec![listener]);
}