use std::net::TcpListener;
use std::sync::{Arc, Mutex};

use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};
use vybe_runtime::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

fn call_import(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-io-contracts-test>");
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

fn bytes_to_vec(value: &Value) -> Vec<u8> {
    let Value::Object(object) = value else {
        return Vec::new();
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(elements) = &object.kind else {
        return Vec::new();
    };
    elements
        .iter()
        .map(|value| value.as_i32().clamp(0, 255) as u8)
        .collect()
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
            std::thread::sleep(std::time::Duration::from_millis(5));
        }
        last
    };

    let server_in = element_at(&accepted, 1);
    let server_out = element_at(&accepted, 2);
    (client_in, client_out, server_in, server_out)
}

#[test]
fn read_on_invalid_handle_returns_null() {
    assert!(matches!(
        call_import("wasi:io/streams", "read", vec![Value::Null, Value::I64(1)]),
        Value::Null
    ));
}

#[test]
fn blocking_read_on_invalid_handle_returns_null() {
    assert!(matches!(
        call_import(
            "wasi:io/streams",
            "blocking-read",
            vec![Value::Null, Value::I64(1)]
        ),
        Value::Null
    ));
}

#[test]
fn skip_on_invalid_handle_returns_null() {
    assert!(matches!(
        call_import("wasi:io/streams", "skip", vec![Value::Null, Value::I64(1)]),
        Value::Null
    ));
}

#[test]
fn blocking_skip_on_invalid_handle_returns_null() {
    assert!(matches!(
        call_import(
            "wasi:io/streams",
            "blocking-skip",
            vec![Value::Null, Value::I64(1)]
        ),
        Value::Null
    ));
}

#[test]
fn subscribe_on_invalid_handle_returns_null() {
    assert!(matches!(
        call_import("wasi:io/streams", "subscribe", vec![Value::Null]),
        Value::Null
    ));
}

#[test]
fn check_write_on_invalid_handle_returns_null() {
    assert!(matches!(
        call_import("wasi:io/streams", "check-write", vec![Value::Null]),
        Value::Null
    ));
}

#[test]
fn write_on_invalid_handle_returns_null() {
    assert!(matches!(
        call_import("wasi:io/streams", "write", vec![Value::Null, bytes(b"hi")]),
        Value::Null
    ));
}

#[test]
fn blocking_write_and_flush_on_invalid_handle_returns_null() {
    assert!(matches!(
        call_import(
            "wasi:io/streams",
            "blocking-write-and-flush",
            vec![Value::Null, bytes(b"hi")]
        ),
        Value::Null
    ));
}

#[test]
fn flush_on_invalid_handle_returns_null() {
    assert!(matches!(
        call_import("wasi:io/streams", "flush", vec![Value::Null]),
        Value::Null
    ));
}

#[test]
fn blocking_flush_on_invalid_handle_returns_null() {
    assert!(matches!(
        call_import("wasi:io/streams", "blocking-flush", vec![Value::Null]),
        Value::Null
    ));
}

#[test]
fn write_zeroes_on_invalid_handle_returns_null() {
    assert!(matches!(
        call_import(
            "wasi:io/streams",
            "write-zeroes",
            vec![Value::Null, Value::I64(4)]
        ),
        Value::Null
    ));
}

#[test]
fn blocking_write_zeroes_and_flush_on_invalid_handle_returns_null() {
    assert!(matches!(
        call_import(
            "wasi:io/streams",
            "blocking-write-zeroes-and-flush",
            vec![Value::Null, Value::I64(4)]
        ),
        Value::Null
    ));
}

#[test]
fn splice_on_invalid_destination_returns_null() {
    let (client_in, _, _, _) = connected_tcp_streams();
    assert!(matches!(
        call_import(
            "wasi:io/streams",
            "splice",
            vec![Value::Null, client_in, Value::I64(4)]
        ),
        Value::Null
    ));
}

#[test]
fn splice_on_invalid_source_returns_null() {
    let (_, _, _, server_out) = connected_tcp_streams();
    assert!(matches!(
        call_import(
            "wasi:io/streams",
            "splice",
            vec![server_out, Value::Null, Value::I64(4)]
        ),
        Value::Null
    ));
}

#[test]
fn blocking_splice_on_invalid_destination_returns_null() {
    let (client_in, _, _, _) = connected_tcp_streams();
    assert!(matches!(
        call_import(
            "wasi:io/streams",
            "blocking-splice",
            vec![Value::Null, client_in, Value::I64(4)]
        ),
        Value::Null
    ));
}

#[test]
fn blocking_splice_on_invalid_source_returns_null() {
    let (_, _, _, server_out) = connected_tcp_streams();
    assert!(matches!(
        call_import(
            "wasi:io/streams",
            "blocking-splice",
            vec![server_out, Value::Null, Value::I64(4)]
        ),
        Value::Null
    ));
}

#[test]
fn poll_ready_on_invalid_pollable_returns_false() {
    assert_eq!(
        call_import("wasi:io/poll", "ready", vec![Value::Null]),
        Value::Bool(false)
    );
}

#[test]
fn poll_block_on_invalid_pollable_returns_null() {
    assert!(matches!(
        call_import("wasi:io/poll", "block", vec![Value::Null]),
        Value::Null
    ));
}

#[test]
fn poll_with_non_array_input_returns_empty_array() {
    let result = call_import("wasi:io/poll", "poll", vec![Value::Null]);
    assert_eq!(bytes_to_vec(&result), Vec::<u8>::new());
}

#[test]
fn read_on_output_stream_returns_null() {
    let (_, _, _, server_out) = connected_tcp_streams();
    assert!(matches!(
        call_import("wasi:io/streams", "read", vec![server_out, Value::I64(4)]),
        Value::Null
    ));
}

#[test]
fn write_on_input_stream_returns_null() {
    let (_, _, server_in, _) = connected_tcp_streams();
    assert!(matches!(
        call_import("wasi:io/streams", "write", vec![server_in, bytes(b"nope")]),
        Value::Null
    ));
}

#[test]
fn flush_on_input_stream_returns_null() {
    let (_, _, server_in, _) = connected_tcp_streams();
    assert!(matches!(
        call_import("wasi:io/streams", "flush", vec![server_in]),
        Value::Null
    ));
}

#[test]
fn skip_beyond_available_bytes_returns_only_available_length() {
    let (_, client_out, server_in, _) = connected_tcp_streams();
    assert!(matches!(
        call_import("wasi:io/streams", "write", vec![client_out, bytes(b"abc")]),
        Value::Null
    ));
    let skipped = call_import("wasi:io/streams", "skip", vec![server_in, Value::I64(10)]);
    assert_eq!(skipped.as_i64(), 3);
}

#[test]
fn blocking_skip_beyond_available_bytes_returns_only_available_length() {
    let (_, client_out, server_in, _) = connected_tcp_streams();
    assert!(matches!(
        call_import("wasi:io/streams", "write", vec![client_out, bytes(b"abc")]),
        Value::Null
    ));
    let skipped = call_import(
        "wasi:io/streams",
        "blocking-skip",
        vec![server_in, Value::I64(10)],
    );
    assert_eq!(skipped.as_i64(), 3);
}
