use std::sync::Arc;
use vybe_bytecode::value::Value;
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn call_import(module: &str, name: &str, pre_stack: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<dotnet-sockets-test>");
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

#[test]
fn dotnet_sockets_tcp_listener_and_client_roundtrip() {
    let listener = call_import("dotnet:sockets", "tcpListenerNew", vec![Value::I32(0)]);
    let port = match &listener {
        Value::Object(obj) => obj.lock().unwrap().properties.get("port").cloned().unwrap_or(Value::Null).as_i32(),
        _ => 0,
    };
    assert!(port > 0, "listener should expose the chosen port");

    let _ = call_import("dotnet:sockets", "tcpListenerStart", vec![listener.clone()]);
    let client = call_import("dotnet:sockets", "tcpClientNew", vec![Value::String(Arc::from("127.0.0.1")), Value::I32(port)]);
    let accepted = {
        let mut last = Value::Null;
        for _ in 0..100 {
            last = call_import("dotnet:sockets", "tcpListenerAccept", vec![listener.clone()]);
            if !matches!(last, Value::Null) {
                break;
            }
            std::thread::sleep(std::time::Duration::from_millis(5));
        }
        last
    };

    let client_stream = call_import("dotnet:sockets", "tcpClientGetStream", vec![client.clone()]);
    let writer = call_import("dotnet:io", "streamWriterNew", vec![client_stream]);
    let reader = call_import("dotnet:io", "streamReaderNew", vec![accepted.clone()]);
    let _ = call_import("dotnet:io", "streamWriterWriteLine", vec![writer, Value::String(Arc::from("ping"))]);
    let line = call_import("dotnet:io", "streamReaderReadLine", vec![reader]);
    assert_eq!(line.as_str(), "ping");

    let _ = call_import("dotnet:sockets", "tcpClientClose", vec![client]);
    let _ = call_import("dotnet:sockets", "tcpClientClose", vec![accepted]);
    let _ = call_import("dotnet:sockets", "tcpListenerStop", vec![listener]);
}