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

fn has_import(module: &str, name: &str) -> bool {
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(module.to_string(), name.to_string()))
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
        values
            .iter()
            .map(|value| Value::I32(*value as i32))
            .collect(),
    ))))
}

fn bytes_to_vec(value: &Value) -> Vec<u8> {
    if let Value::Object(object) = value {
        let object = object.lock().unwrap();
        if let ObjectKind::Array(elements) = &object.kind {
            return elements
                .iter()
                .map(|value| value.as_i32().clamp(0, 255) as u8)
                .collect();
        }
    }
    Vec::new()
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
    assert_eq!(
        len_of(&accepted),
        3,
        "listener should accept the connecting client"
    );

    let server_in = element_at(&accepted, 1);
    let server_out = element_at(&accepted, 2);
    (client_in, client_out, server_in, server_out)
}

#[test]
fn wasi_io_streams_roundtrip_over_tcp_socket_resources() {
    let (client_in, client_out, server_in, server_out) = connected_tcp_streams();

    let incoming_pollable = call_import("wasi:io/streams", "subscribe", vec![server_in.clone()]);
    let payload = bytes(b"ping");
    assert!(matches!(
        call_import(
            "wasi:io/streams",
            "write",
            vec![client_out.clone(), payload]
        ),
        Value::Null
    ));

    let ready = call_import(
        "wasi:io/poll",
        "poll",
        vec![Value::Object(Arc::new(Mutex::new(Object::new_array(
            vec![incoming_pollable],
        ))))],
    );
    assert_eq!(bytes_to_vec(&ready), vec![0]);

    let server_bytes = call_import(
        "wasi:io/streams",
        "blocking-read",
        vec![server_in, Value::I64(4)],
    );
    assert_eq!(bytes_to_vec(&server_bytes), b"ping");

    let reply = bytes(b"pong");
    assert_eq!(
        call_import("wasi:io/streams", "check-write", vec![server_out.clone()]).as_i64(),
        65536
    );
    assert!(matches!(
        call_import(
            "wasi:io/streams",
            "blocking-write-and-flush",
            vec![server_out, reply]
        ),
        Value::Null
    ));

    let client_bytes = call_import(
        "wasi:io/streams",
        "blocking-read",
        vec![client_in, Value::I64(4)],
    );
    assert_eq!(bytes_to_vec(&client_bytes), b"pong");
}

#[test]
fn wasi_io_streams_check_write_returns_null_for_input_stream() {
    let (client_in, _, _, _) = connected_tcp_streams();
    assert!(matches!(
        call_import("wasi:io/streams", "check-write", vec![client_in]),
        Value::Null
    ));
}

#[test]
fn wasi_io_streams_read_zero_length_returns_empty_array() {
    let (_, _, server_in, _) = connected_tcp_streams();
    let result = call_import("wasi:io/streams", "read", vec![server_in, Value::I64(0)]);
    assert_eq!(bytes_to_vec(&result), Vec::<u8>::new());
}

#[test]
fn wasi_io_streams_blocking_read_zero_length_returns_empty_array() {
    let (_, _, server_in, _) = connected_tcp_streams();
    let result = call_import(
        "wasi:io/streams",
        "blocking-read",
        vec![server_in, Value::I64(0)],
    );
    assert_eq!(bytes_to_vec(&result), Vec::<u8>::new());
}

#[test]
fn wasi_io_poll_ready_transitions_after_write() {
    let (_, client_out, server_in, _) = connected_tcp_streams();
    let pollable = call_import("wasi:io/streams", "subscribe", vec![server_in.clone()]);

    assert_eq!(
        call_import("wasi:io/poll", "ready", vec![pollable.clone()]),
        Value::Bool(false)
    );
    assert!(matches!(
        call_import("wasi:io/streams", "write", vec![client_out, bytes(b"xy")]),
        Value::Null
    ));
    assert!(matches!(
        call_import("wasi:io/poll", "block", vec![pollable.clone()]),
        Value::Null
    ));
    assert_eq!(
        call_import("wasi:io/poll", "ready", vec![pollable]),
        Value::Bool(true)
    );
}

#[test]
fn wasi_io_streams_skip_discards_requested_prefix() {
    let (_, client_out, server_in, _) = connected_tcp_streams();
    assert!(matches!(
        call_import(
            "wasi:io/streams",
            "write",
            vec![client_out, bytes(b"pingpong")]
        ),
        Value::Null
    ));

    let skipped = call_import(
        "wasi:io/streams",
        "skip",
        vec![server_in.clone(), Value::I64(4)],
    );
    assert_eq!(skipped.as_i64(), 4);

    let remaining = call_import(
        "wasi:io/streams",
        "blocking-read",
        vec![server_in, Value::I64(8)],
    );
    assert_eq!(bytes_to_vec(&remaining), b"pong");
}

#[test]
fn wasi_io_streams_skip_zero_length_returns_zero() {
    let (_, _, server_in, _) = connected_tcp_streams();
    let skipped = call_import("wasi:io/streams", "skip", vec![server_in, Value::I64(0)]);
    assert_eq!(skipped.as_i64(), 0);
}

#[test]
fn wasi_io_streams_blocking_skip_discards_requested_prefix() {
    let (_, client_out, server_in, _) = connected_tcp_streams();
    assert!(matches!(
        call_import(
            "wasi:io/streams",
            "write",
            vec![client_out, bytes(b"abcdef")]
        ),
        Value::Null
    ));

    let skipped = call_import(
        "wasi:io/streams",
        "blocking-skip",
        vec![server_in.clone(), Value::I64(2)],
    );
    assert_eq!(skipped.as_i64(), 2);

    let remaining = call_import(
        "wasi:io/streams",
        "blocking-read",
        vec![server_in, Value::I64(8)],
    );
    assert_eq!(bytes_to_vec(&remaining), b"cdef");
}

#[test]
fn wasi_io_streams_blocking_skip_zero_length_returns_zero() {
    let (_, _, server_in, _) = connected_tcp_streams();
    let skipped = call_import(
        "wasi:io/streams",
        "blocking-skip",
        vec![server_in, Value::I64(0)],
    );
    assert_eq!(skipped.as_i64(), 0);
}

#[test]
fn wasi_io_streams_write_zeroes_emits_requested_length() {
    let (client_in, _, _, server_out) = connected_tcp_streams();

    assert!(matches!(
        call_import(
            "wasi:io/streams",
            "write-zeroes",
            vec![server_out, Value::I64(3)]
        ),
        Value::Null
    ));
    let bytes = call_import(
        "wasi:io/streams",
        "blocking-read",
        vec![client_in, Value::I64(3)],
    );
    assert_eq!(bytes_to_vec(&bytes), vec![0, 0, 0]);
}

#[test]
fn wasi_io_streams_flush_returns_null_for_output_stream() {
    let (_, _, _, server_out) = connected_tcp_streams();
    assert!(matches!(
        call_import("wasi:io/streams", "flush", vec![server_out]),
        Value::Null
    ));
}

#[test]
fn wasi_io_streams_blocking_flush_returns_null_for_output_stream() {
    let (_, _, _, server_out) = connected_tcp_streams();
    assert!(matches!(
        call_import("wasi:io/streams", "blocking-flush", vec![server_out]),
        Value::Null
    ));
}

#[test]
fn wasi_io_streams_blocking_write_zeroes_and_flush_emits_requested_length() {
    let (client_in, _, _, server_out) = connected_tcp_streams();

    assert!(matches!(
        call_import(
            "wasi:io/streams",
            "blocking-write-zeroes-and-flush",
            vec![server_out, Value::I64(4)]
        ),
        Value::Null
    ));
    let bytes = call_import(
        "wasi:io/streams",
        "blocking-read",
        vec![client_in, Value::I64(4)],
    );
    assert_eq!(bytes_to_vec(&bytes), vec![0, 0, 0, 0]);
}

#[test]
fn wasi_io_streams_splice_moves_bytes_between_connections() {
    let (_, source_out, source_in, _) = connected_tcp_streams();
    let (target_in, _, _, target_out) = connected_tcp_streams();

    assert!(matches!(
        call_import("wasi:io/streams", "write", vec![source_out, bytes(b"copy")]),
        Value::Null
    ));

    let moved = call_import(
        "wasi:io/streams",
        "splice",
        vec![target_out, source_in, Value::I64(4)],
    );
    assert_eq!(moved.as_i64(), 4);

    let copied = call_import(
        "wasi:io/streams",
        "blocking-read",
        vec![target_in, Value::I64(4)],
    );
    assert_eq!(bytes_to_vec(&copied), b"copy");
}

#[test]
fn wasi_io_streams_blocking_splice_moves_bytes_between_connections() {
    let (_, source_out, source_in, _) = connected_tcp_streams();
    let (target_in, _, _, target_out) = connected_tcp_streams();

    assert!(matches!(
        call_import("wasi:io/streams", "write", vec![source_out, bytes(b"pipe")]),
        Value::Null
    ));

    let moved = call_import(
        "wasi:io/streams",
        "blocking-splice",
        vec![target_out, source_in, Value::I64(4)],
    );
    assert_eq!(moved.as_i64(), 4);

    let copied = call_import(
        "wasi:io/streams",
        "blocking-read",
        vec![target_in, Value::I64(4)],
    );
    assert_eq!(bytes_to_vec(&copied), b"pipe");
}

#[test]
fn wasi_io_poll_empty_list_returns_empty_array() {
    let empty = Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))));
    let result = call_import("wasi:io/poll", "poll", vec![empty]);
    assert_eq!(bytes_to_vec(&result), Vec::<u8>::new());
}

#[test]
fn wasi_io_poll_marks_only_ready_pollables() {
    let (_, first_out, first_in, _) = connected_tcp_streams();
    let (_, second_out, second_in, _) = connected_tcp_streams();
    let first_pollable = call_import("wasi:io/streams", "subscribe", vec![first_in]);
    let second_pollable = call_import("wasi:io/streams", "subscribe", vec![second_in]);

    assert!(matches!(
        call_import("wasi:io/streams", "write", vec![second_out, bytes(b"go")]),
        Value::Null
    ));
    let ready = call_import(
        "wasi:io/poll",
        "poll",
        vec![Value::Object(Arc::new(Mutex::new(Object::new_array(
            vec![first_pollable, second_pollable],
        ))))],
    );
    assert_eq!(bytes_to_vec(&ready), vec![1]);

    let _ = first_out;
}

#[test]
fn proposal_io_error_surface_is_registered() {
    let expected = ["[method]error.to-debug-string"];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi:io/error", name))
        .collect::<Vec<_>>();
    assert!(missing.is_empty(), "missing io error imports: {missing:?}");
}

#[test]
fn proposal_io_poll_surface_is_registered() {
    let expected = ["[method]pollable.ready", "[method]pollable.block", "poll"];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi:io/poll", name))
        .collect::<Vec<_>>();
    assert!(missing.is_empty(), "missing io poll imports: {missing:?}");
}

#[test]
fn proposal_io_streams_surface_is_registered() {
    let expected = [
        "[method]input-stream.read",
        "[method]input-stream.blocking-read",
        "[method]input-stream.skip",
        "[method]input-stream.blocking-skip",
        "[method]input-stream.subscribe",
        "[method]output-stream.check-write",
        "[method]output-stream.write",
        "[method]output-stream.blocking-write-and-flush",
        "[method]output-stream.flush",
        "[method]output-stream.blocking-flush",
        "[method]output-stream.subscribe",
        "[method]output-stream.write-zeroes",
        "[method]output-stream.blocking-write-zeroes-and-flush",
        "[method]output-stream.splice",
        "[method]output-stream.blocking-splice",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi:io/streams", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing io streams imports: {missing:?}"
    );
}
