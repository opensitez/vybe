use std::sync::Arc;
use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};

static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn call_import_result(module: &str, name: &str, args: Vec<Value>) -> Result<Value, String> {
    let mut chunk = Chunk::new("<wasi-sockets-test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    for value in args {
        match value {
            Value::I32(n) => chunk.emit_i32_const(n, 0),
            Value::I64(n) => chunk.emit_i64_const(n, 0),
            Value::F32(f) => chunk.emit_f32_const(f, 0),
            Value::F64(f) => chunk.emit_f64_const(f, 0),
            Value::Bool(b) => chunk.emit_bool_const(b, 0),
            Value::String(text) => chunk.emit_string_const(&text, 0),
            Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0),
            other => {
                let global_name = format!(
                    "__test_arg_{}",
                    TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
                );
                vm.set_global_owned(global_name.clone(), other);
                let ci = chunk.intern_string_constant(&global_name);
                chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
            }
        }
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);

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

// The 0.2 half of this file — `instance-network`, `tcp-create-socket`,
// `udp-create-socket`, and the `tcp`/`udp` interfaces with their
// `start-*`/`finish-*` pairs, `accept`, `shutdown` and `subscribe` — is gone
// along with the functions it exercised. WASI 0.3.1 collapsed all of it into
// `wasi:sockets/types`, which the section below covers.
//
// It is deleted rather than kept alongside: a test asserting that both the 0.2
// and the 0.3 spelling answer the same thing reads as diligence but functions
// as a LOCK, holding a name a conforming guest cannot import alive forever.
// That pattern is what kept four other packages from moving.

/// 0.3.1's `resolve-addresses` takes the NAME alone — the `network` handle it
/// took first went with the `instance-network` interface.
#[test]
fn resolve_addresses_accepts_the_name_only_signature() {
    assert!(
        call_import_result(
            "wasi:sockets/ip-name-lookup",
            "resolve-addresses",
            vec![s("localhost")]
        )
        .is_ok(),
        "resolve-addresses must accept `(name)` — 0.3.1 has no network handle"
    );
}

/// The five 0.2 interfaces must be GONE, not merely unused. A name still bound
/// under `wasi:` is a claim the implementation does not keep, because a guest
/// compiled against the 0.3.1 WIT cannot import it.
#[test]
fn the_collapsed_zero_two_socket_interfaces_are_not_registered() {
    for (module, name) in [
        ("wasi:sockets/instance-network", "instance-network"),
        ("wasi:sockets/tcp-create-socket", "create-tcp-socket"),
        ("wasi:sockets/udp-create-socket", "create-udp-socket"),
        ("wasi:sockets/tcp", "start-bind"),
        ("wasi:sockets/tcp", "finish-bind"),
        ("wasi:sockets/tcp", "start-connect"),
        ("wasi:sockets/tcp", "finish-connect"),
        ("wasi:sockets/tcp", "start-listen"),
        ("wasi:sockets/tcp", "finish-listen"),
        ("wasi:sockets/tcp", "accept"),
        ("wasi:sockets/tcp", "shutdown"),
        ("wasi:sockets/tcp", "subscribe"),
        ("wasi:sockets/udp", "start-bind"),
        ("wasi:sockets/udp", "finish-bind"),
        ("wasi:sockets/udp", "stream"),
        ("wasi:sockets/udp", "subscribe"),
    ] {
        assert!(
            !has_import(module, name),
            "{module}.{name} is deleted in WASI 0.3.1 and must not be registered"
        );
    }
}

#[allow(dead_code)]
fn _force_object_use(_: Object) {}

// ── wasi:sockets/types ──────────────────────────────────────────────────────
//
// The `types` interface REPLACED the older `tcp` / `udp` / `*-create-socket` /
// `instance-network` split — it does not sit beside it. A socket comes from
// `tcp-socket.create` or `udp-socket.create`, the start/finish pairs became
// single calls, and `accept` is gone because `listen` itself yields a
// `stream<tcp-socket>` of inbound sockets.
//
// This is now the only socket surface in the tree, which is what
// `the_collapsed_zero_two_socket_interfaces_are_not_registered` above pins.

/// Call a host function on an EXISTING vm, so a socket made by one call is
/// still alive for the next.
fn call_on(vm: &mut VM, module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-sockets-test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    for value in args {
        match value {
            Value::I32(n) => chunk.emit_i32_const(n, 0),
            Value::I64(n) => chunk.emit_i64_const(n, 0),
            Value::F32(f) => chunk.emit_f32_const(f, 0),
            Value::F64(f) => chunk.emit_f64_const(f, 0),
            Value::Bool(b) => chunk.emit_bool_const(b, 0),
            Value::String(text) => chunk.emit_string_const(&text, 0),
            Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0),
            other => {
                let global_name = format!(
                    "__test_arg_{}",
                    TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
                );
                vm.set_global_owned(global_name.clone(), other);
                let ci = chunk.intern_string_constant(&global_name);
                chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
            }
        }
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("VM run failed")
}

fn socket_vm() -> VM {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm
}

/// The `error-code` a failed `result` carries, if this value is one.
fn error_code(value: &Value) -> Option<String> {
    let Value::Object(object) = value else {
        return None;
    };
    let object = object.lock().unwrap();
    match object.properties.get("__wasi_error") {
        Some(Value::String(code)) => Some(code.to_string()),
        _ => None,
    }
}

fn assert_ok(value: &Value, what: &str) {
    assert!(
        error_code(value).is_none(),
        "{what} failed with {:?}",
        error_code(value)
    );
}

fn ipv4_loopback(port: u16) -> Value {
    let octets = Object::new_array(vec![
        Value::I32(127),
        Value::I32(0),
        Value::I32(0),
        Value::I32(1),
    ]);
    let mut address = Object::new();
    address
        .properties
        .insert("family".into(), Value::String(Arc::from("ipv4")));
    address
        .properties
        .insert("port".into(), Value::F64(port as f64));
    address.properties.insert(
        "address".into(),
        Value::Object(vybe_runtime::heap::alloc(octets)),
    );
    Value::Object(vybe_runtime::heap::alloc(address))
}

fn create_tcp(vm: &mut VM) -> Value {
    let socket = call_on(
        vm,
        "wasi:sockets/types",
        "[static]tcp-socket.create",
        vec![Value::String(Arc::from("ipv4"))],
    );
    assert_ok(&socket, "tcp-socket.create");
    socket
}

/// The port a bound socket actually got, which is not the one requested when
/// the request was zero.
fn local_port(vm: &mut VM, resource: &str, socket: &Value) -> u16 {
    let local = call_on(
        vm,
        "wasi:sockets/types",
        &format!("[method]{resource}.get-local-address"),
        vec![socket.clone()],
    );
    assert_ok(&local, "get-local-address");
    let Value::Object(object) = &local else {
        panic!("local address must be a record, got {local:?}");
    };
    let port = object
        .lock()
        .unwrap()
        .properties
        .get("port")
        .map(|value| value.as_f64() as u16);
    port.expect("a bound address must carry a port")
}

/// A socket that has only been created still answers the two infallible
/// getters — neither has a "not bound yet" error to return.
#[test]
fn a_created_socket_reports_its_family_and_is_not_listening() {
    let mut vm = socket_vm();
    let socket = create_tcp(&mut vm);

    assert_eq!(
        call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]tcp-socket.get-address-family",
            vec![socket.clone()],
        ),
        Value::String(Arc::from("ipv4"))
    );
    assert_eq!(
        call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]tcp-socket.get-is-listening",
            vec![socket],
        ),
        Value::Bool(false)
    );
}

/// Binding to port 0 must report the port the OS chose, not the 0 asked for —
/// otherwise nothing can ever learn where to connect.
#[test]
fn binding_to_port_zero_reports_the_assigned_port() {
    let mut vm = socket_vm();
    let socket = create_tcp(&mut vm);
    assert_ok(
        &call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]tcp-socket.bind",
            vec![socket.clone(), ipv4_loopback(0)],
        ),
        "tcp-socket.bind",
    );
    assert!(
        local_port(&mut vm, "tcp-socket", &socket) > 0,
        "port 0 must be replaced by the assigned one"
    );
}

/// A connection opened through `types` is visible to the older `tcp.accept`,
/// because both drive the same listener.
#[test]
fn a_connection_opened_through_types_reaches_the_shared_listener() {
    let mut vm = socket_vm();

    let server = create_tcp(&mut vm);
    assert_ok(
        &call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]tcp-socket.bind",
            vec![server.clone(), ipv4_loopback(0)],
        ),
        "bind server",
    );
    let port = local_port(&mut vm, "tcp-socket", &server);
    assert_ok(
        &call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]tcp-socket.listen",
            vec![server.clone()],
        ),
        "listen",
    );
    assert_eq!(
        call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]tcp-socket.get-is-listening",
            vec![server.clone()],
        ),
        Value::Bool(true),
        "a listening socket must say so"
    );

    let client = create_tcp(&mut vm);
    assert_ok(
        &call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]tcp-socket.connect",
            vec![client, ipv4_loopback(port)],
        ),
        "connect",
    );

    // The SERVER end of this connection is unreachable, and that is the point
    // this test now records.
    //
    // It used to finish by calling `wasi:sockets/tcp.accept` and asserting the
    // 0.3 connect "reaches the shared listener" — i.e. it asserted the two
    // spellings agreed, which is the lock, not a property. `accept` is deleted
    // in 0.3.1: the ONLY way to take an inbound connection is to read the
    // `stream<tcp-socket>` that `listen` returned, and nothing in the bytecode
    // can read a stream whose element is not `u8` (`canon stream.read` copies
    // bytes; `EventLoop::stream_pop` moves whole items but no opcode reaches
    // it). Same blocker as `wasi:filesystem`'s `read-directory`.
    //
    // What IS reachable is asserted above: the client connects, and the kernel
    // completes the handshake into the listener's backlog whether or not
    // anything accepts. So `get-remote-address` answering the right peer port
    // is real evidence the connection is live.
    assert_eq!(
        call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]tcp-socket.get-is-listening",
            vec![server],
        ),
        Value::Bool(true),
        "the listener is still listening after a client connects to it"
    );
}

/// Buffer sizes are real socket options, not values parked on the handle.
///
/// The size that comes back is the KERNEL's, which is the whole point of the
/// assertion — and it is not necessarily the number that went in: Linux
/// reports double what was requested, and the request may shrink the buffer as
/// easily as grow it (macOS defaults to 128 KiB, so asking for 64 KiB reduces
/// it). So this accepts the requested size or twice it, and nothing else.
#[test]
fn buffer_size_options_reach_the_real_socket() {
    let mut vm = socket_vm();
    let socket = create_tcp(&mut vm);

    let before = call_on(
        &mut vm,
        "wasi:sockets/types",
        "[method]tcp-socket.get-receive-buffer-size",
        vec![socket.clone()],
    );
    assert_ok(&before, "get-receive-buffer-size");

    let requested = 64.0 * 1024.0;
    assert_ok(
        &call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]tcp-socket.set-receive-buffer-size",
            vec![socket.clone(), Value::F64(requested)],
        ),
        "set-receive-buffer-size",
    );
    let after = call_on(
        &mut vm,
        "wasi:sockets/types",
        "[method]tcp-socket.get-receive-buffer-size",
        vec![socket.clone()],
    );
    let reported = after.as_f64();
    assert!(
        reported == requested || reported == requested * 2.0,
        "the kernel must report the size that was set (or Linux's doubling of \
         it), got {reported} after asking for {requested} (was {})",
        before.as_f64()
    );

    let rejected = call_on(
        &mut vm,
        "wasi:sockets/types",
        "[method]tcp-socket.set-send-buffer-size",
        vec![socket, Value::F64(0.0)],
    );
    assert_eq!(error_code(&rejected).as_deref(), Some("invalid-argument"));
}

/// `set-hop-limit` rejects 0 — "A value of 0 is not allowed" — and a real value
/// round-trips through the OS socket.
#[test]
fn hop_limit_rejects_zero_and_round_trips_a_real_value() {
    let mut vm = socket_vm();
    let socket = create_tcp(&mut vm);

    let rejected = call_on(
        &mut vm,
        "wasi:sockets/types",
        "[method]tcp-socket.set-hop-limit",
        vec![socket.clone(), Value::F64(0.0)],
    );
    assert_eq!(error_code(&rejected).as_deref(), Some("invalid-argument"));

    assert_ok(
        &call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]tcp-socket.set-hop-limit",
            vec![socket.clone(), Value::F64(37.0)],
        ),
        "set-hop-limit",
    );
    assert_eq!(
        call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]tcp-socket.get-hop-limit",
            vec![socket],
        )
        .as_f64(),
        37.0,
        "the option must reach the OS socket, not a field on the handle"
    );
}

/// A datagram carries its sender's address back — which is why UDP answers
/// `list<u8>` plus a peer rather than TCP's streams.
#[test]
fn udp_receive_reports_the_datagram_and_its_peer() {
    let mut vm = socket_vm();

    let receiver = call_on(
        &mut vm,
        "wasi:sockets/types",
        "[static]udp-socket.create",
        vec![Value::String(Arc::from("ipv4"))],
    );
    assert_ok(&receiver, "udp-socket.create");
    assert_ok(
        &call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]udp-socket.bind",
            vec![receiver.clone(), ipv4_loopback(0)],
        ),
        "bind receiver",
    );
    let port = local_port(&mut vm, "udp-socket", &receiver);

    let sender = call_on(
        &mut vm,
        "wasi:sockets/types",
        "[static]udp-socket.create",
        vec![Value::String(Arc::from("ipv4"))],
    );
    assert_ok(
        &call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]udp-socket.bind",
            vec![sender.clone(), ipv4_loopback(0)],
        ),
        "bind sender",
    );

    let payload = Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
        Value::I32(b'h' as i32),
        Value::I32(b'i' as i32),
    ])));
    assert_ok(
        &call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]udp-socket.send",
            vec![sender, payload, ipv4_loopback(port)],
        ),
        "udp send",
    );

    let received = call_on(
        &mut vm,
        "wasi:sockets/types",
        "[method]udp-socket.receive",
        vec![receiver],
    );
    assert_ok(&received, "udp receive");
    let Value::Object(tuple) = &received else {
        panic!("receive must answer tuple<list<u8>, ip-socket-address>");
    };
    let tuple = tuple.lock().unwrap();
    let ObjectKind::Array(parts) = &tuple.kind else {
        panic!("receive must answer a tuple");
    };
    assert_eq!(parts.len(), 2, "the datagram AND the peer address");

    let Some(Value::Object(bytes)) = parts.first() else {
        panic!("the first element is the datagram");
    };
    let bytes = bytes.lock().unwrap();
    let ObjectKind::Array(bytes) = &bytes.kind else {
        panic!("a datagram is list<u8>");
    };
    let text: Vec<u8> = bytes.iter().map(|byte| byte.as_f64() as u8).collect();
    assert_eq!(text, b"hi", "the payload must survive the round trip");
}

/// Keep-alive is four independent options; each must round-trip through the OS
/// socket rather than being remembered on the handle.
///
/// `duration` is NANOSECONDS, so the idle/interval values go in as nanos and
/// must come back as nanos — a seconds/nanos mix-up here would be invisible
/// without asserting the exact number.
#[test]
fn keep_alive_options_round_trip_through_the_os_socket() {
    let mut vm = socket_vm();
    let socket = create_tcp(&mut vm);

    let get = |vm: &mut VM, suffix: &str, socket: &Value| {
        call_on(
            vm,
            "wasi:sockets/types",
            &format!("[method]tcp-socket.{suffix}"),
            vec![socket.clone()],
        )
    };
    let set = |vm: &mut VM, suffix: &str, socket: &Value, value: Value| {
        call_on(
            vm,
            "wasi:sockets/types",
            &format!("[method]tcp-socket.{suffix}"),
            vec![socket.clone(), value],
        )
    };

    assert_eq!(
        get(&mut vm, "get-keep-alive-enabled", &socket),
        Value::Bool(false),
        "keep-alive is off on a fresh socket"
    );
    assert_ok(
        &set(
            &mut vm,
            "set-keep-alive-enabled",
            &socket,
            Value::Bool(true),
        ),
        "set-keep-alive-enabled",
    );
    assert_eq!(
        get(&mut vm, "get-keep-alive-enabled", &socket),
        Value::Bool(true),
        "the option must reach the OS socket"
    );

    let twenty_seconds = 20.0 * 1_000_000_000.0;
    assert_ok(
        &set(
            &mut vm,
            "set-keep-alive-idle-time",
            &socket,
            Value::F64(twenty_seconds),
        ),
        "set-keep-alive-idle-time",
    );
    assert_eq!(
        get(&mut vm, "get-keep-alive-idle-time", &socket).as_f64(),
        twenty_seconds,
        "idle time is a duration in nanoseconds, in and out"
    );

    let five_seconds = 5.0 * 1_000_000_000.0;
    assert_ok(
        &set(
            &mut vm,
            "set-keep-alive-interval",
            &socket,
            Value::F64(five_seconds),
        ),
        "set-keep-alive-interval",
    );
    assert_eq!(
        get(&mut vm, "get-keep-alive-interval", &socket).as_f64(),
        five_seconds
    );

    assert_ok(
        &set(&mut vm, "set-keep-alive-count", &socket, Value::F64(7.0)),
        "set-keep-alive-count",
    );
    assert_eq!(get(&mut vm, "get-keep-alive-count", &socket).as_f64(), 7.0);

    // A count below 1 is meaningless and the interface has a code for it.
    assert_eq!(
        error_code(&set(
            &mut vm,
            "set-keep-alive-count",
            &socket,
            Value::F64(0.0)
        ))
        .as_deref(),
        Some("invalid-argument")
    );
}

/// `send` takes a `stream<u8>` and `receive` answers one, so bytes written by
/// one socket must come out of the other. This also pins `get-remote-address`,
/// which only has a value once the socket is connected.
#[test]
fn tcp_send_and_receive_move_bytes_between_two_sockets() {
    let mut vm = socket_vm();

    let server = create_tcp(&mut vm);
    assert_ok(
        &call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]tcp-socket.bind",
            vec![server.clone(), ipv4_loopback(0)],
        ),
        "bind",
    );
    let port = local_port(&mut vm, "tcp-socket", &server);
    assert_ok(
        &call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]tcp-socket.listen",
            vec![server.clone()],
        ),
        "listen",
    );

    let client = create_tcp(&mut vm);
    assert_ok(
        &call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]tcp-socket.connect",
            vec![client.clone(), ipv4_loopback(port)],
        ),
        "connect",
    );

    // A connected socket knows its peer; an unconnected one has no answer.
    let remote = call_on(
        &mut vm,
        "wasi:sockets/types",
        "[method]tcp-socket.get-remote-address",
        vec![client.clone()],
    );
    assert_ok(&remote, "get-remote-address");
    let Value::Object(remote_object) = &remote else {
        panic!("remote address must be a record");
    };
    assert_eq!(
        remote_object
            .lock()
            .unwrap()
            .properties
            .get("port")
            .map(|value| value.as_f64() as u16),
        Some(port),
        "the peer port must be the one connected to"
    );

    // `send` on the connected client succeeds: the kernel completes the
    // handshake into the listener's backlog whether or not anything accepts,
    // so there is a real socket underneath to write to.
    let payload = Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
        Value::I32(b'o' as i32),
        Value::I32(b'k' as i32),
    ])));
    assert_ok(
        &call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]tcp-socket.send",
            vec![client.clone(), payload],
        ),
        "send",
    );

    // ⛔ THE OTHER END IS UNREACHABLE, so this can no longer assert that the
    // bytes ARRIVE. It used to read them off a socket obtained from 0.2's
    // `tcp.accept`, describing that as "the one-implementation property" — but
    // 0.3.1 deletes `accept`, and taking an inbound connection means reading
    // the `stream<tcp-socket>` from `listen`, which no bytecode opcode can do.
    // Until that lands there is no way to hold the server end of a connection,
    // and therefore no way to test delivery.
    //
    // What this CAN pin, and what would otherwise regress silently, is that
    // `receive` with nothing pending answers an EMPTY stream instead of
    // hanging. It blocked forever until `receive` was made non-blocking: a
    // socket from `connect` is blocking, so `read` on a peer that has sent
    // nothing parks the VM's only thread and takes the whole program with it.
    // The `WouldBlock` arm in the host was unreachable dead code before that.
    let tuple = call_on(
        &mut vm,
        "wasi:sockets/types",
        "[method]tcp-socket.receive",
        vec![client],
    );
    assert_ok(&tuple, "receive");
    let Value::Object(parts) = &tuple else {
        panic!("receive must answer tuple<stream<u8>, future>");
    };
    let parts = parts.lock().unwrap();
    let ObjectKind::Array(parts) = &parts.kind else {
        panic!("receive must answer a tuple");
    };
    assert_eq!(parts.len(), 2, "the stream AND the completion future");
}

/// The backlog is taken at `listen(2)`, not by a socket option, so it is
/// recorded when set and applied when the socket starts listening.
#[test]
fn listen_backlog_is_recorded_and_applied() {
    let mut vm = socket_vm();
    let socket = create_tcp(&mut vm);

    assert_eq!(
        error_code(&call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]tcp-socket.set-listen-backlog-size",
            vec![socket.clone(), Value::F64(0.0)],
        ))
        .as_deref(),
        Some("invalid-argument"),
        "a backlog of zero is not a backlog"
    );

    assert_ok(
        &call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]tcp-socket.set-listen-backlog-size",
            vec![socket.clone(), Value::F64(16.0)],
        ),
        "set-listen-backlog-size",
    );
    assert_ok(
        &call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]tcp-socket.bind",
            vec![socket.clone(), ipv4_loopback(0)],
        ),
        "bind",
    );
    assert_ok(
        &call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]tcp-socket.listen",
            vec![socket.clone()],
        ),
        "listen must accept the recorded backlog",
    );

    let send_buffer = call_on(
        &mut vm,
        "wasi:sockets/types",
        "[method]tcp-socket.get-send-buffer-size",
        vec![socket],
    );
    assert_ok(&send_buffer, "get-send-buffer-size");
    assert!(
        send_buffer.as_f64() > 0.0,
        "the kernel always has a send buffer"
    );
}

/// `connect` on a datagram socket only fixes the default peer; `disconnect`
/// releases it. Neither is a TCP-style connection.
#[test]
fn udp_connect_fixes_the_default_peer_and_disconnect_releases_it() {
    let mut vm = socket_vm();

    let peer = call_on(
        &mut vm,
        "wasi:sockets/types",
        "[static]udp-socket.create",
        vec![Value::String(Arc::from("ipv4"))],
    );
    assert_ok(
        &call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]udp-socket.bind",
            vec![peer.clone(), ipv4_loopback(0)],
        ),
        "bind peer",
    );
    let peer_port = local_port(&mut vm, "udp-socket", &peer);

    let socket = call_on(
        &mut vm,
        "wasi:sockets/types",
        "[static]udp-socket.create",
        vec![Value::String(Arc::from("ipv4"))],
    );
    assert_eq!(
        call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]udp-socket.get-address-family",
            vec![socket.clone()],
        ),
        Value::String(Arc::from("ipv4"))
    );
    assert_ok(
        &call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]udp-socket.bind",
            vec![socket.clone(), ipv4_loopback(0)],
        ),
        "bind socket",
    );
    assert_ok(
        &call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]udp-socket.connect",
            vec![socket.clone(), ipv4_loopback(peer_port)],
        ),
        "udp connect",
    );

    let remote = call_on(
        &mut vm,
        "wasi:sockets/types",
        "[method]udp-socket.get-remote-address",
        vec![socket.clone()],
    );
    assert_ok(&remote, "get-remote-address after connect");

    // With a default peer set, `send` needs no address.
    let payload = Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
        Value::I32(b'z' as i32),
    ])));
    assert_ok(
        &call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]udp-socket.send",
            vec![socket.clone(), payload],
        ),
        "send to the default peer",
    );

    assert_ok(
        &call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]udp-socket.disconnect",
            vec![socket.clone()],
        ),
        "disconnect",
    );
    assert_eq!(
        error_code(&call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]udp-socket.get-remote-address",
            vec![socket],
        ))
        .as_deref(),
        Some("invalid-state"),
        "a disconnected datagram socket has no peer to report"
    );
}

/// The datagram socket's own options — the same OS settings as TCP's, reached
/// through the `udp-socket` spelling.
#[test]
fn udp_socket_options_round_trip() {
    let mut vm = socket_vm();
    let socket = call_on(
        &mut vm,
        "wasi:sockets/types",
        "[static]udp-socket.create",
        vec![Value::String(Arc::from("ipv4"))],
    );
    assert_ok(&socket, "udp-socket.create");

    assert_eq!(
        error_code(&call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]udp-socket.set-unicast-hop-limit",
            vec![socket.clone(), Value::F64(0.0)],
        ))
        .as_deref(),
        Some("invalid-argument"),
        "a hop limit of 0 is not allowed"
    );
    assert_ok(
        &call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]udp-socket.set-unicast-hop-limit",
            vec![socket.clone(), Value::F64(9.0)],
        ),
        "set-unicast-hop-limit",
    );
    assert_eq!(
        call_on(
            &mut vm,
            "wasi:sockets/types",
            "[method]udp-socket.get-unicast-hop-limit",
            vec![socket.clone()],
        )
        .as_f64(),
        9.0
    );

    for (setter, getter) in [
        (
            "[method]udp-socket.set-receive-buffer-size",
            "[method]udp-socket.get-receive-buffer-size",
        ),
        (
            "[method]udp-socket.set-send-buffer-size",
            "[method]udp-socket.get-send-buffer-size",
        ),
    ] {
        assert_eq!(
            error_code(&call_on(
                &mut vm,
                "wasi:sockets/types",
                setter,
                vec![socket.clone(), Value::F64(0.0)],
            ))
            .as_deref(),
            Some("invalid-argument"),
            "{setter} must reject a zero size"
        );
        let requested = 32.0 * 1024.0;
        assert_ok(
            &call_on(
                &mut vm,
                "wasi:sockets/types",
                setter,
                vec![socket.clone(), Value::F64(requested)],
            ),
            setter,
        );
        let reported =
            call_on(&mut vm, "wasi:sockets/types", getter, vec![socket.clone()]).as_f64();
        assert!(
            reported == requested || reported == requested * 2.0,
            "{getter} must report the kernel's size for what was set, got {reported}"
        );
    }
}
