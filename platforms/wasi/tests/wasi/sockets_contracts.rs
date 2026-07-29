use std::net::TcpListener;
use std::sync::Arc;

use vybe_runtime::value::{ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};
use vybe_runtime::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

fn call_import(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-sockets-contracts-test>");
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

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
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

fn free_port() -> u16 {
    let listener = TcpListener::bind("127.0.0.1:0").expect("bind ephemeral port");
    listener.local_addr().expect("local addr").port()
}

fn listen_on_localhost(port: u16) -> Value {
    let family = s("ipv4");
    let addr = s(&format!("127.0.0.1:{}", port));
    let network = call_import("wasi:sockets/instance-network", "instance-network", vec![]);
    let listener = call_import(
        "wasi:sockets/tcp-create-socket",
        "create-tcp-socket",
        vec![family],
    );
    assert_eq!(
        call_import(
            "wasi:sockets/tcp",
            "start-bind",
            vec![listener.clone(), network, addr]
        ),
        Value::Bool(true)
    );
    assert_eq!(
        call_import("wasi:sockets/tcp", "finish-bind", vec![listener.clone()]),
        Value::Bool(true)
    );
    assert_eq!(
        call_import("wasi:sockets/tcp", "start-listen", vec![listener.clone()]),
        Value::Bool(true)
    );
    assert_eq!(
        call_import("wasi:sockets/tcp", "finish-listen", vec![listener.clone()]),
        Value::Bool(true)
    );
    listener
}

#[test]
fn tcp_is_listening_is_false_before_listen() {
    let socket = call_import(
        "wasi:sockets/tcp-create-socket",
        "create-tcp-socket",
        vec![s("ipv4")],
    );
    assert_eq!(
        call_import("wasi:sockets/tcp", "is-listening", vec![socket]),
        Value::Bool(false)
    );
}

#[test]
fn tcp_address_family_on_invalid_handle_returns_null() {
    assert!(matches!(
        call_import("wasi:sockets/tcp", "address-family", vec![Value::Null]),
        Value::Null
    ));
}

#[test]
fn tcp_local_address_on_invalid_handle_returns_null() {
    assert!(matches!(
        call_import("wasi:sockets/tcp", "local-address", vec![Value::Null]),
        Value::Null
    ));
}

#[test]
fn tcp_remote_address_on_invalid_handle_returns_null() {
    assert!(matches!(
        call_import("wasi:sockets/tcp", "remote-address", vec![Value::Null]),
        Value::Null
    ));
}

#[test]
fn tcp_start_bind_requires_socket_handle() {
    let network = call_import("wasi:sockets/instance-network", "instance-network", vec![]);
    assert!(matches!(
        call_import(
            "wasi:sockets/tcp",
            "start-bind",
            vec![Value::Null, network, s("127.0.0.1:0")]
        ),
        Value::Null
    ));
}

#[test]
fn tcp_start_bind_with_missing_address_returns_null() {
    let socket = call_import(
        "wasi:sockets/tcp-create-socket",
        "create-tcp-socket",
        vec![s("ipv4")],
    );
    let network = call_import("wasi:sockets/instance-network", "instance-network", vec![]);
    assert!(matches!(
        call_import("wasi:sockets/tcp", "start-bind", vec![socket, network]),
        Value::Null
    ));
}

#[test]
fn tcp_finish_bind_returns_true_without_prior_bind() {
    assert_eq!(
        call_import("wasi:sockets/tcp", "finish-bind", vec![Value::Null]),
        Value::Bool(true)
    );
}

#[test]
fn tcp_start_listen_requires_socket_handle() {
    assert!(matches!(
        call_import("wasi:sockets/tcp", "start-listen", vec![Value::Null]),
        Value::Null
    ));
}

#[test]
fn tcp_finish_listen_returns_true_without_prior_listen() {
    assert_eq!(
        call_import("wasi:sockets/tcp", "finish-listen", vec![Value::Null]),
        Value::Bool(true)
    );
}

#[test]
fn tcp_start_connect_requires_socket_handle() {
    let network = call_import("wasi:sockets/instance-network", "instance-network", vec![]);
    assert!(matches!(
        call_import(
            "wasi:sockets/tcp",
            "start-connect",
            vec![Value::Null, network, s("127.0.0.1:1")]
        ),
        Value::Null
    ));
}

#[test]
fn tcp_start_connect_with_missing_address_returns_null() {
    let socket = call_import(
        "wasi:sockets/tcp-create-socket",
        "create-tcp-socket",
        vec![s("ipv4")],
    );
    let network = call_import("wasi:sockets/instance-network", "instance-network", vec![]);
    assert!(matches!(
        call_import("wasi:sockets/tcp", "start-connect", vec![socket, network]),
        Value::Null
    ));
}

#[test]
fn tcp_accept_requires_listener_handle() {
    assert!(matches!(
        call_import("wasi:sockets/tcp", "accept", vec![Value::Null]),
        Value::Null
    ));
}

#[test]
fn tcp_set_listen_backlog_requires_socket_handle() {
    assert!(matches!(
        call_import(
            "wasi:sockets/tcp",
            "set-listen-backlog-size",
            vec![Value::Null, Value::I64(32)]
        ),
        Value::Null
    ));
}

#[test]
fn tcp_subscribe_returns_null_without_argument() {
    assert!(matches!(
        call_import("wasi:sockets/tcp", "subscribe", vec![]),
        Value::Null
    ));
}

#[test]
fn tcp_shutdown_default_mode_returns_true_for_connected_socket() {
    let port = free_port();
    let listener = listen_on_localhost(port);
    let network = call_import("wasi:sockets/instance-network", "instance-network", vec![]);
    let client = call_import(
        "wasi:sockets/tcp-create-socket",
        "create-tcp-socket",
        vec![s("ipv4")],
    );
    assert_eq!(
        call_import(
            "wasi:sockets/tcp",
            "start-connect",
            vec![client.clone(), network, s(&format!("127.0.0.1:{}", port))]
        ),
        Value::Bool(true)
    );
    let streams = call_import("wasi:sockets/tcp", "finish-connect", vec![client.clone()]);
    assert_eq!(len_of(&streams), 2);
    assert_eq!(
        call_import("wasi:sockets/tcp", "shutdown", vec![client]),
        Value::Bool(true)
    );
    let _ = call_import("wasi:sockets/tcp", "accept", vec![listener]);
}

#[test]
fn udp_address_family_on_invalid_handle_returns_null() {
    assert!(matches!(
        call_import("wasi:sockets/udp", "address-family", vec![Value::Null]),
        Value::Null
    ));
}

#[test]
fn udp_local_address_on_invalid_handle_returns_null() {
    assert!(matches!(
        call_import("wasi:sockets/udp", "local-address", vec![Value::Null]),
        Value::Null
    ));
}

#[test]
fn udp_remote_address_on_invalid_handle_returns_null() {
    assert!(matches!(
        call_import("wasi:sockets/udp", "remote-address", vec![Value::Null]),
        Value::Null
    ));
}

#[test]
fn udp_start_bind_requires_socket_handle() {
    let network = call_import("wasi:sockets/instance-network", "instance-network", vec![]);
    assert!(matches!(
        call_import(
            "wasi:sockets/udp",
            "start-bind",
            vec![Value::Null, network, s("127.0.0.1:0")]
        ),
        Value::Null
    ));
}

#[test]
fn udp_start_bind_with_missing_address_returns_null() {
    let socket = call_import(
        "wasi:sockets/udp-create-socket",
        "create-udp-socket",
        vec![s("ipv4")],
    );
    let network = call_import("wasi:sockets/instance-network", "instance-network", vec![]);
    assert!(matches!(
        call_import("wasi:sockets/udp", "start-bind", vec![socket, network]),
        Value::Null
    ));
}

#[test]
fn udp_finish_bind_returns_true_without_prior_bind() {
    assert_eq!(
        call_import("wasi:sockets/udp", "finish-bind", vec![Value::Null]),
        Value::Bool(true)
    );
}

#[test]
fn udp_stream_requires_socket_handle() {
    assert!(matches!(
        call_import("wasi:sockets/udp", "stream", vec![Value::Null, Value::Null]),
        Value::Null
    ));
}

#[test]
fn udp_stream_with_null_remote_preserves_null_remote_address() {
    let port = free_port();
    let socket = call_import(
        "wasi:sockets/udp-create-socket",
        "create-udp-socket",
        vec![s("ipv4")],
    );
    let network = call_import("wasi:sockets/instance-network", "instance-network", vec![]);
    assert_eq!(
        call_import(
            "wasi:sockets/udp",
            "start-bind",
            vec![socket.clone(), network, s(&format!("127.0.0.1:{}", port))]
        ),
        Value::Bool(true)
    );
    let streams = call_import(
        "wasi:sockets/udp",
        "stream",
        vec![socket.clone(), Value::Null],
    );
    assert_eq!(len_of(&streams), 2);
    assert!(matches!(
        call_import("wasi:sockets/udp", "remote-address", vec![socket]),
        Value::Null
    ));
}

#[test]
fn udp_subscribe_returns_null_without_argument() {
    assert!(matches!(
        call_import("wasi:sockets/udp", "subscribe", vec![]),
        Value::Null
    ));
}

#[test]
fn resolve_addresses_returns_stream_for_unknown_host() {
    let network = call_import("wasi:sockets/instance-network", "instance-network", vec![]);
    assert!(matches!(
        call_import(
            "wasi:sockets/ip-name-lookup",
            "resolve-addresses",
            vec![network, s("definitely.invalid.vybe.test")]
        ),
        Value::Object(_)
    ));
}
