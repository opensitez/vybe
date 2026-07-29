use std::net::TcpListener;
use std::sync::Arc;
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};
use vybe_runtime::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

fn call_import(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-sockets-test>");
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

fn call_import_result(module: &str, name: &str, args: Vec<Value>) -> Result<Value, String> {
    let mut chunk = Chunk::new("<wasi-sockets-test>");
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
    vm.run(vec![chunk]).map_err(|error| error.message)
}

fn has_import(module: &str, name: &str) -> bool {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(module.to_string(), name.to_string()))
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
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

fn free_port() -> u16 {
    let listener = TcpListener::bind("127.0.0.1:0").expect("bind ephemeral port");
    listener.local_addr().expect("local addr").port()
}

#[test]
fn instance_network_returns_resource_handle() {
    assert!(matches!(
        call_import("wasi:sockets/instance-network", "instance-network", vec![]),
        Value::Object(_)
    ));
}

fn listen_on_localhost(port: u16) -> (Value, Value, Value) {
    let family = s("ipv4");
    let addr = s(&format!("127.0.0.1:{}", port));
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
            vec![listener.clone(), network, addr],
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

    (listener, family, s(&format!("127.0.0.1:{}", port)))
}

#[test]
fn tcp_create_socket_reports_requested_address_family() {
    let socket = call_import(
        "wasi:sockets/tcp-create-socket",
        "create-tcp-socket",
        vec![s("ipv4")],
    );

    assert_eq!(
        call_import("wasi:sockets/tcp", "address-family", vec![socket]),
        s("ipv4")
    );
}

#[test]
fn tcp_create_socket_defaults_to_ipv4_family() {
    let socket = call_import(
        "wasi:sockets/tcp-create-socket",
        "create-tcp-socket",
        vec![],
    );
    assert_eq!(
        call_import("wasi:sockets/tcp", "address-family", vec![socket]),
        s("ipv4")
    );
}

#[test]
fn tcp_local_address_is_null_before_bind() {
    let socket = call_import(
        "wasi:sockets/tcp-create-socket",
        "create-tcp-socket",
        vec![s("ipv4")],
    );
    assert!(matches!(
        call_import("wasi:sockets/tcp", "local-address", vec![socket]),
        Value::Null
    ));
}

#[test]
fn tcp_bind_and_listen_expose_surface_state() {
    let port = free_port();
    let (listener, _, _) = listen_on_localhost(port);

    assert_eq!(
        call_import("wasi:sockets/tcp", "is-listening", vec![listener.clone()]),
        Value::Bool(true)
    );
    assert!(matches!(
        call_import("wasi:sockets/tcp", "local-address", vec![listener]),
        Value::Object(_)
    ));
}

#[test]
fn tcp_accept_returns_null_before_client_connects() {
    let port = free_port();
    let (listener, _, _) = listen_on_localhost(port);
    assert!(matches!(
        call_import("wasi:sockets/tcp", "accept", vec![listener]),
        Value::Null
    ));
}

#[test]
fn tcp_set_listen_backlog_size_returns_true() {
    let port = free_port();
    let (listener, _, _) = listen_on_localhost(port);
    assert_eq!(
        call_import(
            "wasi:sockets/tcp",
            "set-listen-backlog-size",
            vec![listener, Value::I64(32)]
        ),
        Value::Bool(true)
    );
}

#[test]
fn tcp_subscribe_preserves_socket_behavior() {
    let port = free_port();
    let (listener, _, _) = listen_on_localhost(port);
    let subscribed = call_import("wasi:sockets/tcp", "subscribe", vec![listener]);
    assert_eq!(
        call_import("wasi:sockets/tcp", "is-listening", vec![subscribed]),
        Value::Bool(true)
    );
}

#[test]
fn tcp_connect_populates_addresses_and_stream_pair() {
    let port = free_port();
    let (listener, family, addr) = listen_on_localhost(port);
    let network = call_import("wasi:sockets/instance-network", "instance-network", vec![]);
    let client = call_import(
        "wasi:sockets/tcp-create-socket",
        "create-tcp-socket",
        vec![family],
    );

    assert_eq!(
        call_import(
            "wasi:sockets/tcp",
            "start-connect",
            vec![client.clone(), network, addr],
        ),
        Value::Bool(true)
    );

    let streams = call_import("wasi:sockets/tcp", "finish-connect", vec![client.clone()]);
    assert_eq!(len_of(&streams), 2);
    assert!(matches!(element_at(&streams, 0), Value::Object(_)));
    assert!(matches!(element_at(&streams, 1), Value::Object(_)));
    assert!(matches!(
        call_import("wasi:sockets/tcp", "local-address", vec![client.clone()]),
        Value::Object(_)
    ));
    assert!(matches!(
        call_import("wasi:sockets/tcp", "remote-address", vec![client]),
        Value::Object(_)
    ));

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
    assert_eq!(len_of(&accepted), 3);
}

#[test]
fn tcp_finish_connect_before_start_returns_null() {
    let socket = call_import(
        "wasi:sockets/tcp-create-socket",
        "create-tcp-socket",
        vec![s("ipv4")],
    );
    assert!(matches!(
        call_import("wasi:sockets/tcp", "finish-connect", vec![socket]),
        Value::Null
    ));
}

#[test]
fn tcp_shutdown_returns_true_for_connected_socket() {
    let port = free_port();
    let (_listener, family, addr) = listen_on_localhost(port);
    let network = call_import("wasi:sockets/instance-network", "instance-network", vec![]);
    let client = call_import(
        "wasi:sockets/tcp-create-socket",
        "create-tcp-socket",
        vec![family],
    );

    assert_eq!(
        call_import(
            "wasi:sockets/tcp",
            "start-connect",
            vec![client.clone(), network, addr],
        ),
        Value::Bool(true)
    );
    let streams = call_import("wasi:sockets/tcp", "finish-connect", vec![client.clone()]);
    assert_eq!(len_of(&streams), 2);
    assert_eq!(
        call_import("wasi:sockets/tcp", "shutdown", vec![client, s("send")]),
        Value::Bool(true)
    );
}

#[test]
fn tcp_shutdown_unconnected_socket_returns_null() {
    let socket = call_import(
        "wasi:sockets/tcp-create-socket",
        "create-tcp-socket",
        vec![s("ipv4")],
    );
    assert!(matches!(
        call_import("wasi:sockets/tcp", "shutdown", vec![socket, s("both")]),
        Value::Null
    ));
}

#[test]
fn udp_create_socket_reports_requested_address_family() {
    let socket = call_import(
        "wasi:sockets/udp-create-socket",
        "create-udp-socket",
        vec![s("ipv4")],
    );

    assert_eq!(
        call_import("wasi:sockets/udp", "address-family", vec![socket]),
        s("ipv4")
    );
}

#[test]
fn udp_create_socket_defaults_to_ipv4_family() {
    let socket = call_import(
        "wasi:sockets/udp-create-socket",
        "create-udp-socket",
        vec![],
    );
    assert_eq!(
        call_import("wasi:sockets/udp", "address-family", vec![socket]),
        s("ipv4")
    );
}

#[test]
fn udp_local_address_is_null_before_bind() {
    let socket = call_import(
        "wasi:sockets/udp-create-socket",
        "create-udp-socket",
        vec![s("ipv4")],
    );
    assert!(matches!(
        call_import("wasi:sockets/udp", "local-address", vec![socket]),
        Value::Null
    ));
}

#[test]
fn ip_name_lookup_resolve_addresses_returns_stream_handle() {
    let network = call_import("wasi:sockets/instance-network", "instance-network", vec![]);
    assert!(matches!(
        call_import(
            "wasi:sockets/ip-name-lookup",
            "resolve-addresses",
            vec![network, s("localhost")]
        ),
        Value::Object(_)
    ));
}

#[test]
fn tcp_start_bind_rejects_invalid_socket_address() {
    let socket = call_import(
        "wasi:sockets/tcp-create-socket",
        "create-tcp-socket",
        vec![s("ipv4")],
    );
    let network = call_import("wasi:sockets/instance-network", "instance-network", vec![]);
    assert!(matches!(
        call_import(
            "wasi:sockets/tcp",
            "start-bind",
            vec![socket, network, s("not-an-address")]
        ),
        Value::Null
    ));
}

#[test]
fn udp_remote_address_is_null_before_stream_assignment() {
    let socket = call_import(
        "wasi:sockets/udp-create-socket",
        "create-udp-socket",
        vec![s("ipv4")],
    );
    assert!(matches!(
        call_import("wasi:sockets/udp", "remote-address", vec![socket]),
        Value::Null
    ));
}

#[test]
fn udp_bind_exposes_local_address() {
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
            vec![socket.clone(), network, s(&format!("127.0.0.1:{}", port))],
        ),
        Value::Bool(true)
    );
    assert_eq!(
        call_import("wasi:sockets/udp", "finish-bind", vec![socket.clone()]),
        Value::Bool(true)
    );
    assert!(matches!(
        call_import("wasi:sockets/udp", "local-address", vec![socket]),
        Value::Object(_)
    ));
}

#[test]
fn udp_stream_returns_datagram_stream_pair() {
    let port = free_port();
    let socket = call_import(
        "wasi:sockets/udp-create-socket",
        "create-udp-socket",
        vec![s("ipv4")],
    );
    let network = call_import("wasi:sockets/instance-network", "instance-network", vec![]);
    let remote = s(&format!("127.0.0.1:{}", port));

    assert_eq!(
        call_import(
            "wasi:sockets/udp",
            "start-bind",
            vec![socket.clone(), network, s("127.0.0.1:0")],
        ),
        Value::Bool(true)
    );
    assert_eq!(
        call_import("wasi:sockets/udp", "finish-bind", vec![socket.clone()]),
        Value::Bool(true)
    );

    let streams = call_import("wasi:sockets/udp", "stream", vec![socket.clone(), remote]);
    assert_eq!(len_of(&streams), 2);
    assert!(matches!(element_at(&streams, 0), Value::Object(_)));
    assert!(matches!(element_at(&streams, 1), Value::Object(_)));
    assert!(matches!(
        call_import("wasi:sockets/udp", "remote-address", vec![socket]),
        Value::String(_) | Value::Object(_)
    ));
}

#[test]
fn udp_subscribe_preserves_socket_behavior() {
    let socket = call_import(
        "wasi:sockets/udp-create-socket",
        "create-udp-socket",
        vec![s("ipv4")],
    );
    let subscribed = call_import("wasi:sockets/udp", "subscribe", vec![socket]);
    assert_eq!(
        call_import("wasi:sockets/udp", "address-family", vec![subscribed]),
        s("ipv4")
    );
}

#[test]
fn proposal_ip_name_lookup_resolve_addresses_accepts_name_only_signature() {
    assert!(
        call_import_result(
            "wasi:sockets/ip-name-lookup",
            "resolve-addresses",
            vec![s("localhost")]
        )
        .is_ok(),
        "wasi:sockets/ip-name-lookup.resolve-addresses should be covered with the proposal signature"
    );
}

#[test]
fn proposal_tcp_create_socket_import_is_registered() {
    assert!(
        has_import("wasi:sockets/tcp-create-socket", "create-tcp-socket"),
        "wasi:sockets/tcp-create-socket.create-tcp-socket should be registered"
    );
}

#[test]
fn proposal_tcp_socket_start_bind_import_is_registered() {
    assert!(
        has_import("wasi:sockets/tcp", "[method]tcp-socket.start-bind"),
        "wasi:sockets/tcp.[method]tcp-socket.start-bind should be registered"
    );
}

#[test]
fn proposal_tcp_socket_start_connect_import_is_registered() {
    assert!(
        has_import("wasi:sockets/tcp", "[method]tcp-socket.start-connect"),
        "wasi:sockets/tcp.[method]tcp-socket.start-connect should be registered"
    );
}

#[test]
fn proposal_udp_create_socket_import_is_registered() {
    assert!(
        has_import("wasi:sockets/udp-create-socket", "create-udp-socket"),
        "wasi:sockets/udp-create-socket.create-udp-socket should be registered"
    );
}

#[allow(dead_code)]
fn _force_object_use(_: Object) {}
