use std::net::TcpListener;
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant};

use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_bytecode::capabilities::Capabilities;
use vybe_compiler::compiler::platforms::register_platforms;

fn call_import(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-io-poll-matrix-test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    for value in args {
        let constant = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, constant, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn len_of(value: &Value) -> usize {
    let Value::Object(object) = value else {
        return 0;
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(elements) = &object.kind else {
        return 0;
    };
    elements.len()
}

fn element_at(value: &Value, index: usize) -> Value {
    let Value::Object(object) = value else {
        return Value::Null;
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(elements) = &object.kind else {
        return Value::Null;
    };
    elements.get(index).cloned().unwrap_or(Value::Null)
}

fn bytes(values: &[u8]) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(
        values
            .iter()
            .map(|value| Value::I32(*value as i32))
            .collect(),
    ))))
}

fn indices(value: &Value) -> Vec<i32> {
    let Value::Object(object) = value else {
        return Vec::new();
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(elements) = &object.kind else {
        return Vec::new();
    };
    elements.iter().map(Value::as_i32).collect()
}

fn poll_list(values: Vec<Value>) -> Value {
    call_import(
        "wasi:io/poll",
        "poll",
        vec![Value::Object(Arc::new(Mutex::new(Object::new_array(
            values,
        ))))],
    )
}

fn free_port() -> u16 {
    let listener = TcpListener::bind("127.0.0.1:0").expect("bind ephemeral port");
    listener.local_addr().expect("local addr").port()
}

fn connected_tcp_streams() -> (Value, Value, Value, Value) {
    let port = free_port();
    let addr = Value::String(Arc::from(format!("127.0.0.1:{}", port).as_str()));
    let family = Value::String(Arc::from("ipv4"));

    let network = call_import("wasi:sockets/instance-network", "instance-network", vec![]);
    let listener = call_import(
        "wasi:sockets/tcp-create-socket",
        "create-tcp-socket",
        vec![family.clone()],
    );
    assert_eq!(
        call_import(
            "wasi:sockets/tcp",
            "start-bind",
            vec![listener.clone(), network.clone(), addr.clone()]
        )
        .as_bool(),
        true
    );
    assert_eq!(
        call_import("wasi:sockets/tcp", "finish-bind", vec![listener.clone()]).as_bool(),
        true
    );
    assert_eq!(
        call_import("wasi:sockets/tcp", "start-listen", vec![listener.clone()]).as_bool(),
        true
    );
    assert_eq!(
        call_import("wasi:sockets/tcp", "finish-listen", vec![listener.clone()]).as_bool(),
        true
    );

    let client = call_import(
        "wasi:sockets/tcp-create-socket",
        "create-tcp-socket",
        vec![family],
    );
    assert_eq!(
        call_import(
            "wasi:sockets/tcp",
            "start-connect",
            vec![client.clone(), network, addr]
        )
        .as_bool(),
        true
    );

    let client_streams = call_import("wasi:sockets/tcp", "finish-connect", vec![client]);
    let client_in = element_at(&client_streams, 0);
    let client_out = element_at(&client_streams, 1);

    let accepted = {
        let mut last = Value::Null;
        for _ in 0..100 {
            last = call_import("wasi:sockets/tcp", "accept", vec![listener.clone()]);
            if len_of(&last) == 3 {
                break;
            }
            std::thread::sleep(Duration::from_millis(5));
        }
        last
    };

    let server_in = element_at(&accepted, 1);
    let server_out = element_at(&accepted, 2);
    (client_in, client_out, server_in, server_out)
}

#[test]
fn output_stream_pollable_is_ready_immediately() {
    let (_, _, _, server_out) = connected_tcp_streams();
    let pollable = call_import("wasi:io/streams", "subscribe", vec![server_out]);
    assert_eq!(
        call_import("wasi:io/poll", "ready", vec![pollable]),
        Value::Bool(true)
    );
}

#[test]
fn poll_single_output_stream_returns_index_zero() {
    let (_, _, _, server_out) = connected_tcp_streams();
    let pollable = call_import("wasi:io/streams", "subscribe", vec![server_out]);
    assert_eq!(indices(&poll_list(vec![pollable])), vec![0]);
}

#[test]
fn poll_mixed_invalid_and_ready_output_reports_only_ready_index() {
    let (_, _, _, server_out) = connected_tcp_streams();
    let pollable = call_import("wasi:io/streams", "subscribe", vec![server_out]);
    assert_eq!(indices(&poll_list(vec![Value::Null, pollable])), vec![1]);
}

#[test]
fn duplicate_output_pollables_on_same_stream_both_report_ready() {
    let (_, _, _, server_out) = connected_tcp_streams();
    let first = call_import("wasi:io/streams", "subscribe", vec![server_out.clone()]);
    let second = call_import("wasi:io/streams", "subscribe", vec![server_out]);
    assert_eq!(indices(&poll_list(vec![first, second])), vec![0, 1]);
}

#[test]
fn input_stream_pollable_starts_not_ready_before_write() {
    let (_, _, server_in, _) = connected_tcp_streams();
    let pollable = call_import("wasi:io/streams", "subscribe", vec![server_in]);
    assert_eq!(
        call_import("wasi:io/poll", "ready", vec![pollable]),
        Value::Bool(false)
    );
}

#[test]
fn poll_input_stream_returns_empty_before_write() {
    let (_, _, server_in, _) = connected_tcp_streams();
    let pollable = call_import("wasi:io/streams", "subscribe", vec![server_in]);
    assert!(indices(&poll_list(vec![pollable])).is_empty());
}

#[test]
fn poll_input_stream_returns_index_after_write() {
    let (_, client_out, server_in, _) = connected_tcp_streams();
    let pollable = call_import("wasi:io/streams", "subscribe", vec![server_in]);
    assert!(matches!(
        call_import("wasi:io/streams", "write", vec![client_out, bytes(b"xy")]),
        Value::Null
    ));
    assert_eq!(indices(&poll_list(vec![pollable])), vec![0]);
}

#[test]
fn mixed_output_and_input_pollables_report_both_ready_indices_after_write() {
    let (_, client_out, server_in, server_out) = connected_tcp_streams();
    let input_pollable = call_import("wasi:io/streams", "subscribe", vec![server_in]);
    let output_pollable = call_import("wasi:io/streams", "subscribe", vec![server_out]);
    assert!(matches!(
        call_import("wasi:io/streams", "write", vec![client_out, bytes(b"xy")]),
        Value::Null
    ));
    assert_eq!(
        indices(&poll_list(vec![output_pollable, input_pollable])),
        vec![0, 1]
    );
}

#[test]
fn block_on_ready_output_pollable_returns_quickly() {
    let (_, _, _, server_out) = connected_tcp_streams();
    let pollable = call_import("wasi:io/streams", "subscribe", vec![server_out]);
    let start = Instant::now();
    assert!(matches!(
        call_import("wasi:io/poll", "block", vec![pollable]),
        Value::Null
    ));
    assert!(start.elapsed() < Duration::from_millis(500));
}

#[test]
fn duplicate_input_pollables_both_report_ready_after_write() {
    let (_, client_out, server_in, _) = connected_tcp_streams();
    let first = call_import("wasi:io/streams", "subscribe", vec![server_in.clone()]);
    let second = call_import("wasi:io/streams", "subscribe", vec![server_in]);
    assert!(matches!(
        call_import("wasi:io/streams", "write", vec![client_out, bytes(b"xy")]),
        Value::Null
    ));
    assert_eq!(indices(&poll_list(vec![first, second])), vec![0, 1]);
}
