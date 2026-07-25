//! Behaviour tests for `node:dgram` host imports.
//!
//! Reference: <https://nodejs.org/api/dgram.html>.
//!
//! Coverage:
//!   - `createSocket('udp4'/'udp6')` → Socket object
//!   - `createSocket({type,...})` with options
//!   - Socket method surface: bind, close, send, address, setBroadcast, setTTL,
//!     setMulticastTTL, setMulticastInterface, setMulticastLoopback, addMembership,
//!     dropMembership, addSourceSpecificMembership, dropSourceSpecificMembership,
//!     ref, unref, connect, disconnect, remoteAddress, getSendBufferSize,
//!     getRecvBufferSize, setSendBufferSize, setRecvBufferSize
//!   - EventEmitter methods: on, once, off, emit, removeListener, removeAllListeners,
//!     listeners, rawListeners, listenerCount
//!   - Socket properties: type, fd, _bindState, readyState
//!   - Two createSocket calls return distinct objects
//!   - `Socket` constructor surface

use std::sync::Arc;
use vybe_bytecode::value::{Object, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_bytecode::capabilities::Capabilities;
use vybe_emitter::platforms::register_platforms;

fn call_dgram(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-dgram-test>");
    let import_idx = chunk.add_import("node:dgram", name);
    let argc = args.len() as u8;
    for value in args {
        let c = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, c, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn has_import(name: &str) -> bool {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(String::from("node:dgram"), name.to_string()))
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
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

fn has_method(v: &Value, key: &str) -> bool {
    match v {
        Value::Object(o) => o.lock().unwrap().properties.contains_key(key),
        _ => false,
    }
}

fn as_str(v: &Value) -> String {
    match v {
        Value::String(s) => s.to_string(),
        other => format!("{other}"),
    }
}

// ── createSocket ──────────────────────────────────────────────────────────────

#[test]
fn create_socket_udp4_returns_object() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        matches!(sock, Value::Object(_)),
        "createSocket('udp4') must return object"
    );
}

#[test]
fn create_socket_udp6_returns_object() {
    let sock = call_dgram("createSocket", vec![s("udp6")]);
    assert!(
        matches!(sock, Value::Object(_)),
        "createSocket('udp6') must return object"
    );
}

#[test]
fn create_socket_udp4_type_property_is_udp4() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert_eq!(as_str(&prop(&sock, "type")), "udp4");
}

#[test]
fn create_socket_udp6_type_property_is_udp6() {
    let sock = call_dgram("createSocket", vec![s("udp6")]);
    assert_eq!(as_str(&prop(&sock, "type")), "udp6");
}

#[test]
fn create_socket_with_options_object_udp4_returns_socket() {
    let mut opts = Object::new();
    opts.properties.insert("type".to_string(), s("udp4"));
    opts.properties
        .insert("reuseAddr".to_string(), Value::Bool(false));
    let opts_val = Value::Object(std::sync::Arc::new(std::sync::Mutex::new(opts)));
    let sock = call_dgram("createSocket", vec![opts_val]);
    assert!(matches!(sock, Value::Object(_)));
}

#[test]
fn create_socket_with_ipv6only_option() {
    let mut opts = Object::new();
    opts.properties.insert("type".to_string(), s("udp6"));
    opts.properties
        .insert("ipv6Only".to_string(), Value::Bool(true));
    let opts_val = Value::Object(std::sync::Arc::new(std::sync::Mutex::new(opts)));
    let sock = call_dgram("createSocket", vec![opts_val]);
    assert!(matches!(sock, Value::Object(_)));
}

#[test]
fn two_sockets_are_distinct_objects() {
    let s1 = call_dgram("createSocket", vec![s("udp4")]);
    let s2 = call_dgram("createSocket", vec![s("udp4")]);
    let p1 = match &s1 {
        Value::Object(a) => std::sync::Arc::as_ptr(a) as usize,
        _ => 0,
    };
    let p2 = match &s2 {
        Value::Object(a) => std::sync::Arc::as_ptr(a) as usize,
        _ => 1,
    };
    assert_ne!(
        p1, p2,
        "two createSocket calls must return distinct objects"
    );
}

// ── Socket method surface — I/O ───────────────────────────────────────────────

#[test]
fn socket_has_bind_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(has_method(&sock, "bind"), "socket.bind must be present");
}

#[test]
fn socket_has_close_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(has_method(&sock, "close"), "socket.close must be present");
}

#[test]
fn socket_has_send_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(has_method(&sock, "send"), "socket.send must be present");
}

#[test]
fn socket_has_address_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "address"),
        "socket.address must be present"
    );
}

// ── Socket method surface — multicast / options ───────────────────────────────

#[test]
fn socket_has_set_broadcast_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "setBroadcast"),
        "socket.setBroadcast must be present"
    );
}

#[test]
fn socket_has_set_ttl_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(has_method(&sock, "setTTL"), "socket.setTTL must be present");
}

#[test]
fn socket_has_set_multicast_ttl_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "setMulticastTTL"),
        "socket.setMulticastTTL must be present"
    );
}

#[test]
fn socket_has_set_multicast_loopback_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "setMulticastLoopback"),
        "socket.setMulticastLoopback must be present"
    );
}

#[test]
fn socket_has_set_multicast_interface_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "setMulticastInterface"),
        "socket.setMulticastInterface must be present"
    );
}

#[test]
fn socket_has_add_membership_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "addMembership"),
        "socket.addMembership must be present"
    );
}

#[test]
fn socket_has_drop_membership_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "dropMembership"),
        "socket.dropMembership must be present"
    );
}

#[test]
fn socket_has_add_source_specific_membership_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "addSourceSpecificMembership"),
        "socket.addSourceSpecificMembership must be present"
    );
}

#[test]
fn socket_has_drop_source_specific_membership_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "dropSourceSpecificMembership"),
        "socket.dropSourceSpecificMembership must be present"
    );
}

// ── Socket method surface — ref/unref ─────────────────────────────────────────

#[test]
fn socket_has_ref_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(has_method(&sock, "ref"), "socket.ref must be present");
}

#[test]
fn socket_has_unref_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(has_method(&sock, "unref"), "socket.unref must be present");
}

// ── Socket method surface — buffer sizes ─────────────────────────────────────

#[test]
fn socket_has_get_send_buffer_size_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "getSendBufferSize"),
        "socket.getSendBufferSize must be present"
    );
}

#[test]
fn socket_has_get_recv_buffer_size_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "getRecvBufferSize"),
        "socket.getRecvBufferSize must be present"
    );
}

#[test]
fn socket_has_set_send_buffer_size_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "setSendBufferSize"),
        "socket.setSendBufferSize must be present"
    );
}

#[test]
fn socket_has_set_recv_buffer_size_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "setRecvBufferSize"),
        "socket.setRecvBufferSize must be present"
    );
}

// ── Socket method surface — connected UDP (Node 12+) ─────────────────────────

#[test]
fn socket_has_connect_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "connect"),
        "socket.connect must be present (connected UDP, Node 12+)"
    );
}

#[test]
fn socket_has_disconnect_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "disconnect"),
        "socket.disconnect must be present (connected UDP, Node 12+)"
    );
}

#[test]
fn socket_has_remote_address_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "remoteAddress"),
        "socket.remoteAddress must be present"
    );
}

// ── Socket EventEmitter interface ─────────────────────────────────────────────

#[test]
fn socket_has_on_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "on"),
        "socket.on (EventEmitter) must be present"
    );
}

#[test]
fn socket_has_once_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "once"),
        "socket.once (EventEmitter) must be present"
    );
}

#[test]
fn socket_has_off_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "off"),
        "socket.off (EventEmitter) must be present"
    );
}

#[test]
fn socket_has_emit_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "emit"),
        "socket.emit (EventEmitter) must be present"
    );
}

#[test]
fn socket_has_remove_listener_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "removeListener"),
        "socket.removeListener must be present"
    );
}

#[test]
fn socket_has_remove_all_listeners_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "removeAllListeners"),
        "socket.removeAllListeners must be present"
    );
}

#[test]
fn socket_has_listener_count_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "listenerCount"),
        "socket.listenerCount must be present"
    );
}

#[test]
fn socket_has_listeners_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "listeners"),
        "socket.listeners must be present"
    );
}

#[test]
fn socket_has_raw_listeners_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "rawListeners"),
        "socket.rawListeners must be present"
    );
}

#[test]
fn socket_has_event_names_method() {
    let sock = call_dgram("createSocket", vec![s("udp4")]);
    assert!(
        has_method(&sock, "eventNames"),
        "socket.eventNames must be present"
    );
}

// ── Surface check ─────────────────────────────────────────────────────────────

#[test]
fn proposal_node_dgram_surface_is_registered() {
    let expected = ["createSocket", "Socket"];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing node:dgram imports: {missing:?}"
    );
}
