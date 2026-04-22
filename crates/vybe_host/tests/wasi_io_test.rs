use std::net::TcpListener;
use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn call_import(module: &str, name: &str, pre_stack: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-io-test>");
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

fn len_of(value: &Value) -> usize {
    if let Value::Object(object) = value {
        let object = object.lock().unwrap();
        if let ObjectKind::Array(elements) = &object.kind {
            return elements.len();
        }
    }
    0
}

fn element_at(value: &Value, index: usize) -> Value {
    if let Value::Object(object) = value {
        let object = object.lock().unwrap();
        if let ObjectKind::Array(elements) = &object.kind {
            return elements.get(index).cloned().unwrap_or(Value::Null);
        }
    }
    Value::Null
}

fn bytes(values: &[u8]) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(
        values.iter().map(|value| Value::I32(*value as i32)).collect()
    ))))
}

fn bytes_to_vec(value: &Value) -> Vec<u8> {
    if let Value::Object(object) = value {
        let object = object.lock().unwrap();
        if let ObjectKind::Array(elements) = &object.kind {
            return elements.iter().map(|value| value.as_i32().clamp(0, 255) as u8).collect();
        }
    }
    Vec::new()
}

fn free_port() -> u16 {
    let listener = TcpListener::bind("127.0.0.1:0").expect("bind ephemeral port");
    listener.local_addr().expect("local addr").port()
}

#[test]
fn wasi_io_streams_roundtrip_over_tcp_socket_resources() {
    let port = free_port();
    let addr = Value::String(Arc::from(format!("127.0.0.1:{}", port).as_str()));
    let family = Value::String(Arc::from("ipv4"));

    let network = call_import("wasi:sockets/instance-network", "instance-network", vec![]);
    let listener = call_import("wasi:sockets/tcp-create-socket", "create-tcp-socket", vec![family.clone()]);
    assert_eq!(call_import("wasi:sockets/tcp", "start-bind", vec![listener.clone(), network.clone(), addr.clone()]).as_bool(), true);
    assert_eq!(call_import("wasi:sockets/tcp", "finish-bind", vec![listener.clone()]).as_bool(), true);
    assert_eq!(call_import("wasi:sockets/tcp", "start-listen", vec![listener.clone()]).as_bool(), true);
    assert_eq!(call_import("wasi:sockets/tcp", "finish-listen", vec![listener.clone()]).as_bool(), true);

    let client = call_import("wasi:sockets/tcp-create-socket", "create-tcp-socket", vec![family]);
    assert_eq!(call_import("wasi:sockets/tcp", "start-connect", vec![client.clone(), network, addr]).as_bool(), true);

    let client_streams = call_import("wasi:sockets/tcp", "finish-connect", vec![client]);
    assert_eq!(len_of(&client_streams), 2);
    let client_in = element_at(&client_streams, 0);
    let client_out = element_at(&client_streams, 1);

    let accepted = {
        let mut last = Value::Null;
        for _ in 0..100 {
            last = call_import("wasi:sockets/tcp", "accept", vec![listener.clone()]);
            if len_of(&last) == 3 {
                break;
            }
            std::thread::sleep(std::time::Duration::from_millis(5));
        }
        last
    };
    assert_eq!(len_of(&accepted), 3, "listener should accept the connecting client");
    let server_in = element_at(&accepted, 1);
    let server_out = element_at(&accepted, 2);

    let incoming_pollable = call_import("wasi:io/streams", "subscribe", vec![server_in.clone()]);
    let payload = bytes(b"ping");
    assert!(matches!(call_import("wasi:io/streams", "write", vec![client_out.clone(), payload]), Value::Null));

    let ready = call_import("wasi:io/poll", "poll", vec![Value::Object(Arc::new(Mutex::new(Object::new_array(vec![incoming_pollable]))))]);
    assert_eq!(bytes_to_vec(&ready), vec![0]);

    let server_bytes = call_import("wasi:io/streams", "blocking-read", vec![server_in, Value::I64(4)]);
    assert_eq!(bytes_to_vec(&server_bytes), b"ping");

    let reply = bytes(b"pong");
    assert_eq!(call_import("wasi:io/streams", "check-write", vec![server_out.clone()]).as_i64(), 65536);
    assert!(matches!(call_import("wasi:io/streams", "blocking-write-and-flush", vec![server_out, reply]), Value::Null));

    let client_bytes = call_import("wasi:io/streams", "blocking-read", vec![client_in, Value::I64(4)]);
    assert_eq!(bytes_to_vec(&client_bytes), b"pong");
}