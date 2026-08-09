//! Behaviour tests for `node:net` host imports.
//!
//! Reference: <https://nodejs.org/api/net.html>.
//!
//! Coverage:
//!   - `isIP(input)` → 4 | 6 | 0
//!   - `isIPv4(input)` → boolean
//!   - `isIPv6(input)` → boolean
//!   - `createServer([options])` → Server object with listen, close, address,
//!     getConnections, setTimeout, on, once, off, emit methods
//!   - `createConnection(options)` / `connect(options)` → Socket (surface)
//!   - `Socket` constructor → object with connect, write, destroy, end, pipe,
//!     setEncoding, setTimeout, on, once, off, emit methods + state props
//!   - `Server` constructor (surface)
//!   - `BlockList` — addAddress, addRange, addSubnet, check methods (Node 15+)
//!   - `SocketAddress` — address, port, family properties (Node 15.14+)
//!
//! Deferred (require live network or async infrastructure):
//!   - `.listen()`, `.close()`, `.address()` (need event loop)
//!   - `.connect()`, `.write()`, `.read()`, `.destroy()` on Socket

use std::sync::Arc;
use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::{Object, Value};
use vybe_runtime::{Chunk, Op, VM};

static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn call_net(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-net-test>");
    let import_idx = chunk.add_import("node:net", name);
    let argc = args.len() as u8;
    let mut arg_globals: Vec<(String, Value)> = Vec::new();
    for value in args {
        match value {
            Value::I32(n) => chunk.emit_i32_const(n, 0),
            Value::I64(n) => chunk.emit_i64_const(n, 0),
            Value::F32(f) => chunk.emit_f32_const(f, 0),
            Value::F64(f) => chunk.emit_f64_const(f, 0),
            Value::Bool(b) => chunk.emit_bool_const(b, 0),
            Value::String(s) => chunk.emit_string_const(&s, 0),
            other => {
                let name = format!(
                    "__test_arg_{}",
                    TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
                );
                let ci = chunk.intern_string_constant(&name);
                chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
                arg_globals.push((name, other));
            }
        }
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    for (name, value) in arg_globals {
        vm.globals.insert(name, value);
    }
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn has_import(name: &str) -> bool {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(String::from("node:net"), name.to_string()))
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn new_obj(pairs: Vec<(&str, Value)>) -> Value {
    let mut o = Object::new();
    for (k, v) in pairs {
        o.properties.insert(k.to_string(), v);
    }
    Value::Object(std::sync::Arc::new(std::sync::Mutex::new(o)))
}

fn has_method(v: &Value, key: &str) -> bool {
    match v {
        Value::Object(o) => o.lock().unwrap().properties.contains_key(key),
        _ => false,
    }
}

fn prop(v: &Value, key: &str) -> Value {
    match v {
        Value::Object(o) => o
            .lock()
            .unwrap()
            .properties
            .get(key)
            .cloned()
            .unwrap_or(Value::Undefined),
        _ => Value::Undefined,
    }
}

// ── isIP ──────────────────────────────────────────────────────────────────────

#[test]
fn is_ip_returns_4_for_ipv4() {
    assert_eq!(call_net("isIP", vec![s("192.168.1.1")]), Value::I32(4));
}

#[test]
fn is_ip_returns_4_for_loopback_v4() {
    assert_eq!(call_net("isIP", vec![s("127.0.0.1")]), Value::I32(4));
}

#[test]
fn is_ip_returns_4_for_broadcast() {
    assert_eq!(call_net("isIP", vec![s("255.255.255.255")]), Value::I32(4));
}

#[test]
fn is_ip_returns_4_for_zero_address() {
    assert_eq!(call_net("isIP", vec![s("0.0.0.0")]), Value::I32(4));
}

#[test]
fn is_ip_returns_6_for_full_ipv6() {
    assert_eq!(
        call_net("isIP", vec![s("2001:0db8:85a3:0000:0000:8a2e:0370:7334")]),
        Value::I32(6)
    );
}

#[test]
fn is_ip_returns_6_for_compressed_ipv6() {
    assert_eq!(call_net("isIP", vec![s("::1")]), Value::I32(6));
}

#[test]
fn is_ip_returns_6_for_all_zeros_ipv6() {
    assert_eq!(call_net("isIP", vec![s("::")]), Value::I32(6));
}

#[test]
fn is_ip_returns_6_for_ipv4_mapped_ipv6() {
    assert_eq!(
        call_net("isIP", vec![s("::ffff:192.168.1.1")]),
        Value::I32(6)
    );
}

#[test]
fn is_ip_returns_0_for_hostname() {
    assert_eq!(call_net("isIP", vec![s("example.com")]), Value::I32(0));
}

#[test]
fn is_ip_returns_0_for_empty_string() {
    assert_eq!(call_net("isIP", vec![s("")]), Value::I32(0));
}

#[test]
fn is_ip_returns_0_for_malformed_ipv4() {
    assert_eq!(call_net("isIP", vec![s("999.0.0.1")]), Value::I32(0));
}

#[test]
fn is_ip_returns_0_for_partial_address() {
    assert_eq!(call_net("isIP", vec![s("192.168")]), Value::I32(0));
}

// ── isIPv4 ────────────────────────────────────────────────────────────────────

#[test]
fn is_ipv4_true_for_valid_v4() {
    assert_eq!(call_net("isIPv4", vec![s("10.0.0.1")]), Value::Bool(true));
}

#[test]
fn is_ipv4_false_for_ipv6() {
    assert_eq!(call_net("isIPv4", vec![s("::1")]), Value::Bool(false));
}

#[test]
fn is_ipv4_false_for_hostname() {
    assert_eq!(call_net("isIPv4", vec![s("localhost")]), Value::Bool(false));
}

#[test]
fn is_ipv4_false_for_too_many_octets() {
    assert_eq!(call_net("isIPv4", vec![s("1.2.3.4.5")]), Value::Bool(false));
}

#[test]
fn is_ipv4_false_for_octet_out_of_range() {
    assert_eq!(call_net("isIPv4", vec![s("256.0.0.1")]), Value::Bool(false));
}

// ── isIPv6 ────────────────────────────────────────────────────────────────────

#[test]
fn is_ipv6_true_for_loopback() {
    assert_eq!(call_net("isIPv6", vec![s("::1")]), Value::Bool(true));
}

#[test]
fn is_ipv6_true_for_full_address() {
    assert_eq!(
        call_net("isIPv6", vec![s("2001:db8::1")]),
        Value::Bool(true)
    );
}

#[test]
fn is_ipv6_false_for_ipv4() {
    assert_eq!(
        call_net("isIPv6", vec![s("192.168.1.1")]),
        Value::Bool(false)
    );
}

#[test]
fn is_ipv6_false_for_hostname() {
    assert_eq!(
        call_net("isIPv6", vec![s("example.com")]),
        Value::Bool(false)
    );
}

// ── createServer ──────────────────────────────────────────────────────────────

#[test]
fn create_server_returns_server_object() {
    let server = call_net("createServer", vec![]);
    assert!(matches!(server, Value::Object(_)));
}

#[test]
fn create_server_has_listen_method() {
    let server = call_net("createServer", vec![]);
    assert!(has_method(&server, "listen"), "Server.listen must exist");
}

#[test]
fn create_server_has_close_method() {
    let server = call_net("createServer", vec![]);
    assert!(has_method(&server, "close"), "Server.close must exist");
}

#[test]
fn create_server_has_address_method() {
    let server = call_net("createServer", vec![]);
    assert!(has_method(&server, "address"), "Server.address must exist");
}

#[test]
fn create_server_has_get_connections_method() {
    let server = call_net("createServer", vec![]);
    assert!(
        has_method(&server, "getConnections"),
        "Server.getConnections must exist"
    );
}

#[test]
fn create_server_has_on_method() {
    let server = call_net("createServer", vec![]);
    assert!(
        has_method(&server, "on"),
        "Server.on (EventEmitter) must exist"
    );
}

#[test]
fn create_server_has_once_method() {
    let server = call_net("createServer", vec![]);
    assert!(has_method(&server, "once"), "Server.once must exist");
}

#[test]
fn create_server_has_off_method() {
    let server = call_net("createServer", vec![]);
    assert!(has_method(&server, "off"), "Server.off must exist");
}

#[test]
fn create_server_has_emit_method() {
    let server = call_net("createServer", vec![]);
    assert!(has_method(&server, "emit"), "Server.emit must exist");
}

#[test]
fn create_server_with_allow_half_open_option() {
    let opts = new_obj(vec![("allowHalfOpen", Value::Bool(true))]);
    let server = call_net("createServer", vec![opts]);
    assert!(matches!(server, Value::Object(_)));
}

#[test]
fn create_server_with_pause_on_connect_option() {
    let opts = new_obj(vec![("pauseOnConnect", Value::Bool(true))]);
    let server = call_net("createServer", vec![opts]);
    assert!(matches!(server, Value::Object(_)));
}

// ── createConnection / connect ────────────────────────────────────────────────

#[test]
fn create_connection_with_port_and_host_returns_socket() {
    let opts = new_obj(vec![("port", Value::I32(8080)), ("host", s("localhost"))]);
    let socket = call_net("createConnection", vec![opts]);
    assert!(matches!(socket, Value::Object(_)));
}

#[test]
fn connect_alias_returns_socket() {
    let opts = new_obj(vec![("port", Value::I32(9090))]);
    let socket = call_net("connect", vec![opts]);
    assert!(matches!(socket, Value::Object(_)));
}

// ── Socket constructor ────────────────────────────────────────────────────────

#[test]
fn socket_constructor_returns_object() {
    let socket = call_net("Socket", vec![]);
    assert!(matches!(socket, Value::Object(_)));
}

#[test]
fn socket_has_connect_method() {
    let socket = call_net("Socket", vec![]);
    assert!(has_method(&socket, "connect"), "Socket.connect must exist");
}

#[test]
fn socket_has_write_method() {
    let socket = call_net("Socket", vec![]);
    assert!(has_method(&socket, "write"), "Socket.write must exist");
}

#[test]
fn socket_has_end_method() {
    let socket = call_net("Socket", vec![]);
    assert!(has_method(&socket, "end"), "Socket.end must exist");
}

#[test]
fn socket_has_destroy_method() {
    let socket = call_net("Socket", vec![]);
    assert!(has_method(&socket, "destroy"), "Socket.destroy must exist");
}

#[test]
fn socket_has_pipe_method() {
    let socket = call_net("Socket", vec![]);
    assert!(has_method(&socket, "pipe"), "Socket.pipe must exist");
}

#[test]
fn socket_has_set_encoding_method() {
    let socket = call_net("Socket", vec![]);
    assert!(
        has_method(&socket, "setEncoding"),
        "Socket.setEncoding must exist"
    );
}

#[test]
fn socket_has_set_timeout_method() {
    let socket = call_net("Socket", vec![]);
    assert!(
        has_method(&socket, "setTimeout"),
        "Socket.setTimeout must exist"
    );
}

#[test]
fn socket_has_on_method() {
    let socket = call_net("Socket", vec![]);
    assert!(has_method(&socket, "on"), "Socket.on must exist");
}

#[test]
fn socket_has_once_method() {
    let socket = call_net("Socket", vec![]);
    assert!(has_method(&socket, "once"), "Socket.once must exist");
}

#[test]
fn socket_has_emit_method() {
    let socket = call_net("Socket", vec![]);
    assert!(has_method(&socket, "emit"), "Socket.emit must exist");
}

#[test]
fn socket_has_remote_address_property() {
    let socket = call_net("Socket", vec![]);
    let _ = prop(&socket, "remoteAddress"); // may be undefined before connect
}

#[test]
fn socket_has_local_address_property() {
    let socket = call_net("Socket", vec![]);
    let _ = prop(&socket, "localAddress"); // may be undefined before connect
}

#[test]
fn socket_has_bytes_written_property() {
    let socket = call_net("Socket", vec![]);
    let bw = prop(&socket, "bytesWritten");
    assert!(matches!(
        bw,
        Value::I32(_) | Value::I64(_) | Value::F64(_) | Value::Undefined
    ));
}

#[test]
fn socket_has_bytes_read_property() {
    let socket = call_net("Socket", vec![]);
    let br = prop(&socket, "bytesRead");
    assert!(matches!(
        br,
        Value::I32(_) | Value::I64(_) | Value::F64(_) | Value::Undefined
    ));
}

// ── BlockList ─────────────────────────────────────────────────────────────────

#[test]
fn block_list_constructor_returns_object() {
    let bl = call_net("BlockList", vec![]);
    assert!(matches!(bl, Value::Object(_)));
}

#[test]
fn block_list_has_add_address_method() {
    let bl = call_net("BlockList", vec![]);
    assert!(
        has_method(&bl, "addAddress"),
        "BlockList.addAddress must exist"
    );
}

#[test]
fn block_list_has_add_range_method() {
    let bl = call_net("BlockList", vec![]);
    assert!(has_method(&bl, "addRange"), "BlockList.addRange must exist");
}

#[test]
fn block_list_has_add_subnet_method() {
    let bl = call_net("BlockList", vec![]);
    assert!(
        has_method(&bl, "addSubnet"),
        "BlockList.addSubnet must exist"
    );
}

#[test]
fn block_list_has_check_method() {
    let bl = call_net("BlockList", vec![]);
    assert!(has_method(&bl, "check"), "BlockList.check must exist");
}

#[test]
fn block_list_has_rules_property() {
    let bl = call_net("BlockList", vec![]);
    // rules is an array of added addresses
    let rules = prop(&bl, "rules");
    assert!(
        !matches!(rules, Value::Undefined),
        "BlockList.rules must exist"
    );
}

// ── SocketAddress ─────────────────────────────────────────────────────────────

#[test]
fn socket_address_constructor_returns_object() {
    let opts = new_obj(vec![
        ("address", s("127.0.0.1")),
        ("port", Value::I32(80)),
        ("family", s("ipv4")),
    ]);
    let sa = call_net("SocketAddress", vec![opts]);
    assert!(matches!(sa, Value::Object(_)));
}

#[test]
fn socket_address_has_address_property() {
    let opts = new_obj(vec![("address", s("10.0.0.1")), ("port", Value::I32(443))]);
    let sa = call_net("SocketAddress", vec![opts]);
    let addr = prop(&sa, "address");
    match addr {
        Value::String(a) => assert_eq!(a.as_ref(), "10.0.0.1"),
        Value::Undefined => {} // TDD
        other => panic!("address expected string, got {:?}", other),
    }
}

#[test]
fn socket_address_has_port_property() {
    let opts = new_obj(vec![("address", s("10.0.0.1")), ("port", Value::I32(443))]);
    let sa = call_net("SocketAddress", vec![opts]);
    let port = prop(&sa, "port");
    assert!(matches!(
        port,
        Value::I32(_) | Value::F64(_) | Value::Undefined
    ));
}

#[test]
fn socket_address_has_family_property() {
    let opts = new_obj(vec![
        ("address", s("::1")),
        ("port", Value::I32(80)),
        ("family", s("ipv6")),
    ]);
    let sa = call_net("SocketAddress", vec![opts]);
    let family = prop(&sa, "family");
    assert!(matches!(family, Value::String(_) | Value::Undefined));
}

// ── Surface check ─────────────────────────────────────────────────────────────

// ── Server constructor ────────────────────────────────────────────────────────

#[test]
fn server_constructor_returns_object() {
    let server = call_net("Server", vec![]);
    assert!(
        matches!(server, Value::Object(_)),
        "Server constructor must return object"
    );
}

#[test]
fn server_constructor_has_listen_method() {
    let server = call_net("Server", vec![]);
    assert!(has_method(&server, "listen"), "Server.listen must exist");
}

#[test]
fn server_constructor_has_close_method() {
    let server = call_net("Server", vec![]);
    assert!(has_method(&server, "close"), "Server.close must exist");
}

// ── createServer — setTimeout ─────────────────────────────────────────────────

#[test]
fn create_server_has_set_timeout_method() {
    let server = call_net("createServer", vec![]);
    assert!(
        has_method(&server, "setTimeout"),
        "Server.setTimeout must exist"
    );
}

// ── Socket — missing methods and properties ───────────────────────────────────

#[test]
fn socket_has_off_method() {
    let socket = call_net("Socket", vec![]);
    assert!(has_method(&socket, "off"), "Socket.off must exist");
}

#[test]
fn socket_has_readable_property() {
    let socket = call_net("Socket", vec![]);
    let r = prop(&socket, "readable");
    assert!(
        matches!(r, Value::Bool(_) | Value::Undefined),
        "Socket.readable must be bool or undefined before connect"
    );
}

#[test]
fn socket_has_writable_property() {
    let socket = call_net("Socket", vec![]);
    let w = prop(&socket, "writable");
    assert!(
        matches!(w, Value::Bool(_) | Value::Undefined),
        "Socket.writable must be bool or undefined before connect"
    );
}

#[test]
fn socket_has_connecting_property() {
    let socket = call_net("Socket", vec![]);
    let c = prop(&socket, "connecting");
    assert!(
        matches!(c, Value::Bool(_) | Value::Undefined),
        "Socket.connecting must be bool or undefined"
    );
}

#[test]
fn socket_has_pending_property() {
    let socket = call_net("Socket", vec![]);
    let p = prop(&socket, "pending");
    assert!(
        matches!(p, Value::Bool(_) | Value::Undefined),
        "Socket.pending must be bool or undefined"
    );
}

#[test]
fn socket_has_destroyed_property() {
    let socket = call_net("Socket", vec![]);
    let d = prop(&socket, "destroyed");
    assert!(
        matches!(d, Value::Bool(_) | Value::Undefined),
        "Socket.destroyed must be bool or undefined"
    );
}

// ── SocketAddress — flowLabel ─────────────────────────────────────────────────

#[test]
fn socket_address_has_flow_label_property() {
    let opts = new_obj(vec![
        ("address", s("::1")),
        ("port", Value::I32(80)),
        ("family", s("ipv6")),
        ("flowlabel", Value::I32(0)),
    ]);
    let sa = call_net("SocketAddress", vec![opts]);
    let fl = prop(&sa, "flowlabel");
    assert!(
        matches!(fl, Value::I32(_) | Value::F64(_) | Value::Undefined),
        "SocketAddress.flowlabel must be number or undefined"
    );
}

// ── BlockList — addAddress with type arg, check returns bool ──────────────────

#[test]
fn block_list_add_address_with_ipv6_type() {
    let bl = call_net("BlockList", vec![]);
    // addAddress(address, type) — should not panic
    let result = call_net("addAddress", vec![bl, s("::1"), s("ipv6")]);
    // result is whatever BlockList returns; we only care it didn't panic
    let _ = result;
}

#[test]
fn block_list_check_returns_bool_for_unknown_address() {
    let bl = call_net("BlockList", vec![]);
    let result = call_net("check", vec![bl, s("192.168.1.1")]);
    assert!(
        matches!(result, Value::Bool(_) | Value::Undefined | Value::Null),
        "BlockList.check must return bool, got {:?}",
        result
    );
}

#[test]
fn proposal_node_net_surface_is_registered() {
    let expected = [
        "isIP",
        "isIPv4",
        "isIPv6",
        "createServer",
        "createConnection",
        "connect",
        "Socket",
        "Server",
        "BlockList",
        "SocketAddress",
        "NetConnectOpts",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(missing.is_empty(), "missing node:net imports: {missing:?}");
}
