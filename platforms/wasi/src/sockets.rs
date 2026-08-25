//! `wasi:sockets@0.3.1` — the `types` interface.
//!
//! 0.3.1 collapsed `network`, `instance-network`, `tcp`, `tcp-create-socket`,
//! `udp` and `udp-create-socket` into a single `types` interface, and deleted
//! `accept`, `shutdown` and every start-*/finish-* pair: `listen` answers a
//! `stream<tcp-socket>` that IS the accept queue, and `send`/`receive` exchange
//! `stream<u8>` rather than bytes.
//!
//! This file previously opened by announcing itself as `System.Net.Sockets`
//! and carried a TcpClient/TcpListener/UdpClient surface from the VB-interpreter
//! era. Those thirty-six functions had no callers and were invisible behind a
//! file-level `#![allow(dead_code)]`; .NET reaches sockets through
//! `platforms/dotnet`, not through a WASI host.


use std::collections::HashMap;
use std::io::{Read, Write};
use std::net::{
    IpAddr, Ipv4Addr, Ipv6Addr, SocketAddr, TcpListener, TcpStream, UdpSocket,
};
use std::sync::{
    Arc, Mutex,
    atomic::{AtomicU64, Ordering},
};
use std::thread;
use std::time::{Duration, Instant};
use vybe_runtime::typedef::{Method, TypeDef};
use vybe_runtime::value::Object;
use vybe_runtime::component::ValType;
use vybe_runtime::{HostContext, VM, Value};

static NEXT_ID: AtomicU64 = AtomicU64::new(1);

pub struct SocketState {
    pub tcp_streams: HashMap<u64, TcpStream>,
    tcp_listeners: HashMap<u64, TcpListener>,
    pending_accepts: HashMap<u64, Vec<TcpStream>>,
    udp_sockets: HashMap<u64, UdpSocket>,
    /// Sockets that exist but have not yet become a listener or a stream.
    ///
    /// WASI 0.3 separates `create` from `bind` from `listen`/`connect`, which
    /// POSIX has always done and `std::net` cannot: `TcpListener::bind` and
    /// `TcpStream::connect` each create AND commit the socket in one call. A
    /// socket lives here from `tcp-socket.create` until it commits, at which
    /// point it moves into `tcp_listeners` or `tcp_streams` under the same id.
    raw_sockets: HashMap<u64, socket2::Socket>,
    /// `stream_id → listener_id` for every live `listen()` stream.
    ///
    /// 0.3.1 has no `accept`: `listen` answers a `stream<tcp-socket>` and each
    /// inbound connection is one element. Connections arrive AFTER `listen`
    /// returns, so the stream has to be filled by something other than the
    /// call that made it — this is the mapping the producer
    /// (`accept-into-stream`) uses to find the listener a reader is waiting on.
    listen_streams: HashMap<u64, u64>,
    /// `stream_id → socket_id` for every live `receive()` stream, and the same
    /// story: bytes arrive after the call that made the stream returned.
    recv_streams: HashMap<u64, u64>,
}

pub fn get_state() -> Arc<Mutex<SocketState>> {
    use std::sync::OnceLock;
    static STATE: OnceLock<Arc<Mutex<SocketState>>> = OnceLock::new();
    STATE
        .get_or_init(|| {
            Arc::new(Mutex::new(SocketState {
                tcp_streams: HashMap::new(),
                tcp_listeners: HashMap::new(),
                pending_accepts: HashMap::new(),
                udp_sockets: HashMap::new(),
                raw_sockets: HashMap::new(),
                listen_streams: HashMap::new(),
                recv_streams: HashMap::new(),
            }))
        })
        .clone()
}

/// VM hot-reset (bucket C/D): close/drop all open sockets so a reused VM never
/// carries a prior run's OS socket handles. See `vmhotresetplan.md`.
pub fn reset() {
    if let Ok(mut s) = get_state().lock() {
        s.tcp_streams.clear();
        s.tcp_listeners.clear();
        s.pending_accepts.clear();
        s.udp_sockets.clear();
        s.raw_sockets.clear();
        s.listen_streams.clear();
        s.recv_streams.clear();
    }
}















pub fn register(vm: &mut VM) {
    // WASI 0.3.1 collapsed `network`, `instance-network`, `tcp`,
    // `tcp-create-socket`, `udp` and `udp-create-socket` into one
    // `wasi:sockets/types` interface, so `register_wasi_sockets_0_3` is the
    // whole of this package's surface: the 0.2 interfaces are not registered
    // at all, because a name a conforming guest cannot import has no business
    // sitting inside the `wasi:` namespace.
    //
    // `register_wasi_io` is NOT called any more. It was the last of the
    register_wasi_sockets_0_3(vm);
    // The inbound-connection producer. Host plumbing, not an interface — see
    // `PRODUCER_MODULE`.
    vm.register_host_fn(
        PRODUCER_MODULE,
        PRODUCER_ACCEPT,
        Box::new(accept_into_stream),
    );
    vm.register_host_fn(
        PRODUCER_MODULE,
        PRODUCER_RECEIVE,
        Box::new(receive_into_stream),
    );
    // LAST: a vtable entry is a host-fn registry INDEX, so every function it
    // names has to exist first.
    register_socket_types(vm);
}

// ── wasi:sockets@0.3.0 — `types` ────────────────────────────────────────────
//
// 0.3 collapses 0.2's `tcp`, `udp`, `tcp-create-socket`, `udp-create-socket`
// and `instance-network` into ONE `types` interface: a socket comes from
// `tcp-socket.create` / `udp-socket.create`, the start/finish pairs become
// single calls, and `accept` disappears because `listen` itself yields a
// `stream<tcp-socket>`.
//
// Nothing here is a second socket stack. The handles are the same objects the
// 0.2 functions take, the OS sockets live in the same `get_state()` maps, and
// addresses go through the same `parse_ip_socket_address` /
// `socket_addr_to_value` helpers — so a socket bound through 0.3 can be
// accepted through 0.2 and vice versa.

/// `error-code` per `proposals/sockets/wit/types.wit`, matching the
/// representation `wasi:http` uses for its own `result` error side.
fn socket_err(code: &str) -> Value {
    let mut object = Object::new();
    object
        .properties
        .insert("__wasi_error".into(), Value::String(Arc::from(code)));
    Value::Object(vybe_runtime::heap::alloc(object))
}

/// Map a std IO error onto the closest `error-code` variant.
fn socket_err_from(error: &std::io::Error) -> Value {
    use std::io::ErrorKind::*;
    socket_err(match error.kind() {
        PermissionDenied => "access-denied",
        AddrInUse => "address-in-use",
        AddrNotAvailable => "address-not-bindable",
        ConnectionRefused => "connection-refused",
        ConnectionReset => "connection-reset",
        ConnectionAborted => "connection-aborted",
        BrokenPipe => "connection-broken",
        TimedOut => "timeout",
        InvalidInput => "invalid-argument",
        Unsupported => "not-supported",
        _ => "other",
    })
}

/// The handle's `__socket_id`, or 0.
fn socket_id_of(handle: &Arc<Mutex<Object>>) -> u64 {
    handle
        .lock()
        .unwrap()
        .properties
        .get("__socket_id")
        .map(|value| value.as_f64() as u64)
        .unwrap_or(0)
}

fn handle_string(handle: &Arc<Mutex<Object>>, key: &str) -> Option<String> {
    match handle.lock().unwrap().properties.get(key) {
        Some(Value::String(text)) => Some(text.to_string()),
        _ => None,
    }
}

fn set_handle(handle: &Arc<Mutex<Object>>, key: &str, value: Value) {
    handle.lock().unwrap().properties.insert(key.into(), value);
}

/// A freshly created, unbound socket resource.
///
/// `get-address-family` and `get-is-listening` return bare values rather than
/// results, so they must answer correctly on a socket that has been created
/// and nothing else — which is why the family is written here and not deferred
/// to `bind`.
/// `TypeRegistry` ids for the two socket resources, filled in by
/// [`register_socket_types`].
///
/// A handle carries its type id so MEMBER DISPATCH can find the method —
/// `namespaceplan.md` §"Three structures, three phases": member dispatch is
/// receiver-based through the `TypeRegistry` vtable and NEVER resolves through
/// the namespace tree; only ctors and statics are reachable by path walk. A
/// socket registered as bare host functions therefore had no way to be called
/// as `sock.bind(...)` by ANY language, which is why each one grew its own
/// socket method table.
static TCP_TYPE_ID: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);
static UDP_TYPE_ID: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

/// The `wasi:sockets/types` methods, as the names a RECEIVER is called with.
///
/// Only the ones that are a single host call. `accept`, `send` and `recv` are
/// deliberately absent: 0.3.1 makes them Component Model STREAM sequences
/// (`canon stream.read` off `listen`'s `stream<tcp-socket>`, `stream.write`
/// into `send`'s `stream<u8>`), and a canon built-in is a guest INSTRUCTION —
/// there is no host function for a vtable to point at. Those stay compiler
/// emits; see `languages/python/src/emitter/socket_adapter.rs`.
const TCP_METHODS: &[(&str, &str)] = &[
    ("bind", "[method]tcp-socket.bind"),
    ("connect", "[method]tcp-socket.connect"),
    ("listen", "[method]tcp-socket.listen"),
    ("getsockname", "[method]tcp-socket.get-local-address"),
    ("getpeername", "[method]tcp-socket.get-remote-address"),
];

const UDP_METHODS: &[(&str, &str)] = &[
    ("bind", "[method]udp-socket.bind"),
    ("connect", "[method]udp-socket.connect"),
    ("disconnect", "[method]udp-socket.disconnect"),
    ("send", "[method]udp-socket.send"),
    ("recv", "[method]udp-socket.receive"),
    ("getsockname", "[method]udp-socket.get-local-address"),
    ("getpeername", "[method]udp-socket.get-remote-address"),
];

/// Give the two socket resources a vtable.
///
/// Called AFTER the host functions are registered — a method is stored as a
/// host-fn REGISTRY INDEX, so the vtable cannot be built until the functions it
/// points at exist. That ordering is the whole reason `finalize` is a separate
/// phase from `init` (`platforms/ecma/src/builtin_types.rs`).
fn register_socket_types(vm: &mut VM) {
    for (resource, methods, slot) in [
        ("tcp-socket", TCP_METHODS, &TCP_TYPE_ID),
        ("udp-socket", UDP_METHODS, &UDP_TYPE_ID),
    ] {
        let mut type_def = TypeDef::new(resource);
        type_def.interface = Some("wasi:sockets/types".into());
        type_def.is_resource = true;
        for (method, host_name) in methods {
            if let Some(idx) = vm
                .host_registry
                .get(&("wasi:sockets/types".to_string(), (*host_name).to_string()))
            {
                type_def
                    .methods
                    .insert((*method).to_string(), Method::HostFn(*idx));
            }
        }
        let type_id = vm.type_registry.register(type_def);
        vm.register_host_resource_type_export("wasi:sockets/types", resource, type_id);
        slot.store(type_id, Ordering::Relaxed);
    }
}

fn new_socket_handle(kind: &str, family: &str) -> Value {
    let id = NEXT_ID.fetch_add(1, Ordering::Relaxed);
    // Typed, not bare: an object with `type_id: 0` reaches member dispatch as
    // an untyped bag and no method is found on it.
    let type_id = if kind == "udp-socket" {
        UDP_TYPE_ID.load(Ordering::Relaxed)
    } else {
        TCP_TYPE_ID.load(Ordering::Relaxed)
    };
    let mut object = Object::new_typed(type_id);
    object
        .properties
        .insert("__type".into(), Value::String(Arc::from(kind)));
    object
        .properties
        .insert("__socket_id".into(), Value::F64(id as f64));
    object
        .properties
        .insert("__address_family".into(), Value::String(Arc::from(family)));
    object
        .properties
        .insert("__state".into(), Value::String(Arc::from("unbound")));
    object
        .properties
        .insert("islistening".into(), Value::Bool(false));
    Value::Object(vybe_runtime::heap::alloc(object))
}

fn family_arg(args: &[Value], index: usize) -> String {
    match args.get(index) {
        Some(Value::String(text)) if text.as_ref() == "ipv6" => "ipv6".to_string(),
        _ => "ipv4".to_string(),
    }
}

fn socket2_domain(family: &str) -> socket2::Domain {
    if family == "ipv6" {
        socket2::Domain::IPV6
    } else {
        socket2::Domain::IPV4
    }
}

/// Resolve a `ip-socket-address` argument to a `SocketAddr`.
fn resolve_socket_addr(value: &Value) -> Option<SocketAddr> {
    let (host, port, _) = parse_ip_socket_address(value)?;
    format!("{host}:{port}").parse::<SocketAddr>().ok()
}

/// Run `body` with the socket handle from `args[0]`, or answer `invalid-state`.
fn with_socket(args: &[Value], body: impl FnOnce(&Arc<Mutex<Object>>, u64) -> Value) -> Value {
    match socket_arg(args) {
        Some(handle) => {
            let id = socket_id_of(&handle);
            body(&handle, id)
        }
        None => socket_err("invalid-state"),
    }
}

/// Apply `body` to the raw (created but uncommitted) socket for `id`.
fn with_raw_socket(id: u64, body: impl FnOnce(&socket2::Socket) -> Value) -> Value {
    let state = get_state();
    let guard = state.lock().unwrap();
    match guard.raw_sockets.get(&id) {
        Some(socket) => body(socket),
        None => socket_err("invalid-state"),
    }
}

/// A socket option lives on the raw socket before it commits and on the
/// connected `TcpStream` afterwards; both are the same OS socket, so an option
/// call has to look in whichever map currently owns it.
fn with_socket2_view(id: u64, body: impl FnOnce(&socket2::Socket) -> Value) -> Value {
    let state = get_state();
    let guard = state.lock().unwrap();
    if let Some(socket) = guard.raw_sockets.get(&id) {
        return body(socket);
    }
    // `SockRef` borrows the fd without taking ownership, so the stream stays
    // usable and is not closed when the view is dropped.
    if let Some(stream) = guard.tcp_streams.get(&id) {
        return body(&socket2::SockRef::from(stream));
    }
    if let Some(listener) = guard.tcp_listeners.get(&id) {
        return body(&socket2::SockRef::from(listener));
    }
    socket_err("invalid-state")
}

/// The resource type id `own<tcp-socket>` handles are minted under.
///
/// Distinct from `udp-socket`'s so a handle of the wrong type is caught by
/// `canon resource.drop`'s type check rather than silently dropping the wrong
/// socket — that check is the entire reason handles carry a type at all.
const TCP_SOCKET_TYPE_ID: u32 = 1;

/// Socket id → the handle object that id names.
///
/// The other half of `create_own_resource`. §`canon resource.rep` fixes a
/// resource's representation as an i32 and nothing else, so a host cannot make
/// the socket OBJECT the representation and expect a guest to get it back —
/// what comes out of `resource.rep` is an integer. The honest representation is
/// therefore the socket's own id, and this is what turns that id back into the
/// handle every method here already takes.
///
/// A separate lock from [`get_state`] on purpose: [`socket_arg`] runs at the
/// top of nearly every registration, and making it reach for the state mutex
/// would deadlock the moment one of them resolved an argument while already
/// holding it.
fn socket_reps() -> &'static Mutex<HashMap<u64, Arc<Mutex<Object>>>> {
    use std::sync::OnceLock;
    static REPS: OnceLock<Mutex<HashMap<u64, Arc<Mutex<Object>>>>> = OnceLock::new();
    REPS.get_or_init(|| Mutex::new(HashMap::new()))
}

/// Record `handle` under its own id so an `own<tcp-socket>` can be resolved
/// back from the i32 a guest holds.
fn remember_socket_rep(handle: &Arc<Mutex<Object>>) -> u64 {
    let id = socket_id_of(handle);
    socket_reps().lock().unwrap().insert(id, handle.clone());
    id
}

/// Where the inbound-connection producer is registered.
///
/// Deliberately NOT under `wasi:`. It is not an interface any guest may
/// import — the WIT declares no such function, and a profile that named it
/// would be inventing a verb exactly the way `fs.rs` once did. It is host
/// plumbing that the VM reaches through the ordinary registry, the same way it
/// reaches `ecma:promise.resolve` internally, and the `wasi:` namespace gate in
/// `interface_coverage.rs` correctly ignores it.
const PRODUCER_MODULE: &str = "wasi-host:sockets";
const PRODUCER_ACCEPT: &str = "accept-into-stream";
const PRODUCER_RECEIVE: &str = "receive-into-stream";

/// How long a parked `accept()` waits for a connection before giving up.
///
/// A bound rather than a forever-block: a listener with no client would
/// otherwise hang the process with no way to say why. On expiry nothing is
/// pushed, the reader parks, and the event loop reports the deadlock BY NAME
/// (`run_event_loop`'s parked-copy check) — which is the honest ending, and
/// the one that a silent empty read never gave.
const ACCEPT_WAIT: Duration = Duration::from_secs(20);

/// Accept one inbound connection for `stream_id`'s listener and push it.
///
/// Called by a reader that is about to park on the stream, so blocking here IS
/// the suspension: the guest asked for a connection and this is the wait.
fn accept_into_stream(ctx: &mut HostContext, args: &[Value]) -> Value {
    let stream_id = args.first().map(|v| v.as_f64() as u64).unwrap_or(0);
    let listener_id = get_state()
        .lock()
        .unwrap()
        .listen_streams
        .get(&stream_id)
        .copied();
    let Some(listener_id) = listener_id else {
        return Value::Null;
    };

    // Poll rather than a blocking accept: the listener is shared state behind
    // a mutex, and holding that lock across a blocking syscall would stop
    // every other socket call in the program.
    let deadline = Instant::now() + ACCEPT_WAIT;
    loop {
        let accepted = {
            let state = get_state();
            let guard = state.lock().unwrap();
            match guard.tcp_listeners.get(&listener_id) {
                Some(listener) => listener.accept().ok(),
                // The listener is gone — the socket was closed under us. Not
                // an error, just nothing more to hand over.
                None => return Value::Null,
            }
        };
        if let Some((stream, peer)) = accepted {
            let child_id = NEXT_ID.fetch_add(1, Ordering::Relaxed);
            let family = if peer.is_ipv6() { "ipv6" } else { "ipv4" };
            let child = new_socket_handle("tcp-socket", family);
            if let Value::Object(object) = &child {
                let mut o = object.lock().unwrap();
                o.properties
                    .insert("__socket_id".into(), Value::F64(child_id as f64));
                o.properties
                    .insert("__state".into(), Value::String(Arc::from("connected")));
                o.properties
                    .insert("__remote_address".into(), socket_addr_to_value(peer));
            }
            get_state()
                .lock()
                .unwrap()
                .tcp_streams
                .insert(child_id, stream);
            // Same representation rule as `listen`'s own push: the id, which
            // `resource.rep` can actually give back, never the object.
            let rep = match &child {
                Value::Object(object) => Value::F64(remember_socket_rep(object) as f64),
                _ => child.clone(),
            };
            let item = ctx
                .create_own_resource(TCP_SOCKET_TYPE_ID, rep)
                .unwrap_or(child);
            ctx.stream_push(stream_id, item);
            return Value::Null;
        }
        if Instant::now() >= deadline {
            return Value::Null;
        }
        thread::sleep(Duration::from_millis(1));
    }
}

/// Read bytes for `stream_id`'s socket and push them, or close on a real
/// end-of-stream.
///
/// The `receive()` counterpart of [`accept_into_stream`]. `read` answering
/// `WouldBlock` means the peer has not sent yet, which is a reason to WAIT;
/// `read` answering zero means the peer has closed, which is the one thing
/// that legitimately ends the stream. Conflating those two is what made an
/// empty socket read as a disconnect.
fn receive_into_stream(ctx: &mut HostContext, args: &[Value]) -> Value {
    let stream_id = args.first().map(|v| v.as_f64() as u64).unwrap_or(0);
    let socket_id = get_state()
        .lock()
        .unwrap()
        .recv_streams
        .get(&stream_id)
        .copied();
    let Some(socket_id) = socket_id else {
        return Value::Null;
    };

    let deadline = Instant::now() + ACCEPT_WAIT;
    loop {
        let mut chunk = [0u8; 65536];
        let outcome = {
            let state = get_state();
            let mut guard = state.lock().unwrap();
            match guard.tcp_streams.get_mut(&socket_id) {
                Some(stream) => {
                    let _ = stream.set_nonblocking(true);
                    Some(stream.read(&mut chunk))
                }
                None => None,
            }
        };
        match outcome {
            // The socket is gone: nothing more can arrive, so this really is
            // the end of the stream.
            None => {
                ctx.stream_close(stream_id);
                return Value::Null;
            }
            Some(Ok(0)) => {
                ctx.stream_close(stream_id);
                return Value::Null;
            }
            Some(Ok(read)) => {
                for byte in &chunk[..read] {
                    ctx.stream_push(stream_id, Value::I32(*byte as i32));
                }
                return Value::Null;
            }
            Some(Err(error))
                if error.kind() == std::io::ErrorKind::WouldBlock
                    || error.kind() == std::io::ErrorKind::TimedOut => {}
            // A genuine error ends the stream — the reader gets `DROPPED`
            // rather than waiting on a socket that cannot recover.
            Some(Err(_)) => {
                ctx.stream_close(stream_id);
                return Value::Null;
            }
        }
        if Instant::now() >= deadline {
            return Value::Null;
        }
        thread::sleep(Duration::from_millis(1));
    }
}

fn register_tcp_socket_0_3(vm: &mut VM) {
    // `create: static func(address-family) -> result<tcp-socket, error-code>`
    // The OS socket exists from here, so options can be set before `bind` —
    // which is the order POSIX requires for `SO_RCVBUF`/`SO_SNDBUF`.
    vm.register_host_fn(
        "wasi:sockets/types",
        "[static]tcp-socket.create",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let family = family_arg(args, 0);
            let socket = match socket2::Socket::new(
                socket2_domain(&family),
                socket2::Type::STREAM,
                Some(socket2::Protocol::TCP),
            ) {
                Ok(socket) => socket,
                Err(error) => return socket_err_from(&error),
            };
            let handle = new_socket_handle("tcp-socket", &family);
            if let Value::Object(object) = &handle {
                let id = socket_id_of(object);
                get_state().lock().unwrap().raw_sockets.insert(id, socket);
            }
            handle
        }),
    );

    // `bind: func(local-address) -> result<_, error-code>`
    vm.register_host_fn(
        "wasi:sockets/types",
        "[method]tcp-socket.bind",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            with_socket(args, |handle, id| {
                let Some(address) = method_arg(args, 0).and_then(resolve_socket_addr) else {
                    return socket_err("invalid-argument");
                };
                let outcome = with_raw_socket(id, |socket| match socket.bind(&address.into()) {
                    Ok(()) => Value::Null,
                    Err(error) => socket_err_from(&error),
                });
                if matches!(outcome, Value::Null) {
                    // "If the port is zero the socket is bound to a random free
                    // port" — so the BOUND address is read back, never the
                    // requested one, or `get-local-address` reports port 0.
                    let bound = with_raw_socket(id, |socket| {
                        socket
                            .local_addr()
                            .ok()
                            .and_then(|addr| addr.as_socket())
                            .map(socket_addr_to_value)
                            .unwrap_or(Value::Null)
                    });
                    set_handle(handle, "__local_address", bound);
                    set_handle(handle, "__state", Value::String(Arc::from("bound")));
                }
                outcome
            })
        }),
    );

    // `connect: async func(remote-address) -> result<_, error-code>`
    //
    // The socket leaves `raw_sockets` and becomes a `TcpStream` under the same
    // id, which is what the 0.2 `tcp.accept` / `io/streams` functions read.
    vm.register_host_fn(
        "wasi:sockets/types",
        "[method]tcp-socket.connect",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            with_socket(args, |handle, id| {
                let Some(address) = method_arg(args, 0).and_then(resolve_socket_addr) else {
                    return socket_err("invalid-argument");
                };
                let state = get_state();
                let mut guard = state.lock().unwrap();
                let Some(socket) = guard.raw_sockets.remove(&id) else {
                    return socket_err("invalid-state");
                };
                if let Err(error) = socket.connect(&address.into()) {
                    guard.raw_sockets.insert(id, socket);
                    return socket_err_from(&error);
                }
                let local = socket
                    .local_addr()
                    .ok()
                    .and_then(|addr| addr.as_socket())
                    .map(socket_addr_to_value);
                guard.tcp_streams.insert(id, TcpStream::from(socket));
                drop(guard);

                set_handle(handle, "__state", Value::String(Arc::from("connected")));
                set_handle(handle, "__remote_address", socket_addr_to_value(address));
                if let Some(local) = local {
                    set_handle(handle, "__local_address", local);
                }
                Value::Null
            })
        }),
    );

    // `listen: func() -> result<stream<tcp-socket>, error-code>`
    //
    // 0.3 has no `accept`: the stream of inbound sockets IS the return value.
    // The VM's `stream` is an eager buffer — values are pushed and the stream
    // is closed — so it cannot express "connections keep arriving". What this
    // returns is therefore the connections already pending when `listen` was
    // called, and the listener stays registered so `wasi:sockets/tcp.accept`
    // continues to hand over later ones. That shortfall is the same
    vm.register_host_fn(
        "wasi:sockets/types",
        "[method]tcp-socket.listen",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            with_socket(args, |handle, id| {
                let backlog = handle
                    .lock()
                    .unwrap()
                    .properties
                    .get("__listen_backlog")
                    .map(|value| value.as_f64() as i32)
                    .unwrap_or(128);
                let state = get_state();
                let mut guard = state.lock().unwrap();
                let Some(socket) = guard.raw_sockets.remove(&id) else {
                    return socket_err("invalid-state");
                };
                if let Err(error) = socket.listen(backlog) {
                    guard.raw_sockets.insert(id, socket);
                    return socket_err_from(&error);
                }
                let listener = TcpListener::from(socket);
                let _ = listener.set_nonblocking(true);
                let mut inbound = Vec::new();
                while let Ok((stream, peer)) = listener.accept() {
                    inbound.push((stream, peer));
                }
                guard.tcp_listeners.insert(id, listener);

                let mut accepted = Vec::new();
                for (stream, peer) in inbound {
                    let child_id = NEXT_ID.fetch_add(1, Ordering::Relaxed);
                    let family = if peer.is_ipv6() { "ipv6" } else { "ipv4" };
                    guard.tcp_streams.insert(child_id, stream);
                    let child = new_socket_handle("tcp-socket", family);
                    if let Value::Object(object) = &child {
                        object
                            .lock()
                            .unwrap()
                            .properties
                            .insert("__socket_id".into(), Value::F64(child_id as f64));
                        object
                            .lock()
                            .unwrap()
                            .properties
                            .insert("__state".into(), Value::String(Arc::from("connected")));
                        object
                            .lock()
                            .unwrap()
                            .properties
                            .insert("__remote_address".into(), socket_addr_to_value(peer));
                    }
                    accepted.push(child);
                }
                drop(guard);

                set_handle(handle, "__state", Value::String(Arc::from("listening")));
                set_handle(handle, "islistening", Value::Bool(true));
                set_handle(handle, "__listener_id", Value::F64(id as f64));

                // `listen: func() -> result<stream<tcp-socket>, error-code>`.
                //
                // Each accepted socket is minted as a real `own<tcp-socket>`
                // handle, so `canon stream.read` can lower it as the i32
                // handle-table index the canonical ABI says a resource is. The
                // socket object itself becomes the handle's REPRESENTATION and
                // stays host-side, which is what `resource.rep` exists to
                // recover.
                //
                // This was untyped until a host could mint a resource at all:
                // `resource.new` is a bytecode built-in, so the object was the
                // only thing a host function could return, and an object
                // lowers as `as_i32()` of nothing.
                let (stream_val, stream_id) =
                    ctx.create_stream_of(Some(ValType::Own("tcp-socket".to_string())));
                for socket in accepted {
                    // The REPRESENTATION is the socket's id, not the object.
                    // §`canon resource.rep` gives back an i32, so an object put
                    // in here comes out as `as_i32()` of nothing — the guest
                    // would hold a handle it could never turn into a socket.
                    // The id is an opaque integer this component chose, which
                    // is exactly what a representation is meant to be, and
                    // `socket_reps` is what resolves it on the way back in.
                    let rep = match &socket {
                        Value::Object(object) => Value::F64(remember_socket_rep(object) as f64),
                        _ => socket.clone(),
                    };
                    // No handle table (bare VM, no component instance) means no
                    // handle to mint — the socket object is then the only thing
                    // there is to hand over, matching `create_stream`'s own
                    // fallback rather than dropping the connection.
                    let item = ctx
                        .create_own_resource(TCP_SOCKET_TYPE_ID, rep)
                        .unwrap_or(socket);
                    ctx.stream_push(stream_id, item);
                }

                // NOT closed. The spec is explicit that this is "a single
                // perpetual stream that should only close on fatal errors",
                // and closing it here made `listen` answer only the
                // connections that happened to be pending BEFORE it was
                // called — reliably none, since a server listens before its
                // clients exist. `accept()` then returned nothing, forever,
                // and reported clean EOF while doing it.
                //
                // What fills it instead is the PRODUCER: a reader that is
                // about to park calls `accept-into-stream`, which blocks on
                // this listener and pushes the connection it gets. That is the
                // whole of the machinery — a host function returns once, so
                // the reader has to be the one who asks.
                get_state()
                    .lock()
                    .unwrap()
                    .listen_streams
                    .insert(stream_id, id);
                ctx.set_stream_producer(stream_id, PRODUCER_MODULE, PRODUCER_ACCEPT);
                stream_val
            })
        }),
    );

    // `send: func(data: stream<u8>) -> future<result<_, error-code>>`
    vm.register_host_fn(
        "wasi:sockets/types",
        "[method]tcp-socket.send",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let handle = socket_arg(args);
            let bytes = match method_arg(args, 0) {
                Some(value) => ctx.stream_drain(value),
                None => Vec::new(),
            };
            let outcome = match handle {
                Some(handle) => {
                    let id = socket_id_of(&handle);
                    let state = get_state();
                    let mut guard = state.lock().unwrap();
                    match guard.tcp_streams.get_mut(&id) {
                        Some(stream) => {
                            match stream.write_all(&bytes).and_then(|()| stream.flush()) {
                                Ok(()) => Value::Null,
                                Err(error) => socket_err_from(&error),
                            }
                        }
                        None => socket_err("invalid-state"),
                    }
                }
                None => socket_err("invalid-state"),
            };
            let (future_val, future_id) = ctx.create_future();
            ctx.resolve_future(future_id, outcome);
            future_val
        }),
    );

    // `receive: func() -> tuple<stream<u8>, future<result<_, error-code>>>`
    vm.register_host_fn(
        "wasi:sockets/types",
        "[method]tcp-socket.receive",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let mut buffer = Vec::new();
            let mut failure = Value::Null;
            if let Some(handle) = socket_arg(args) {
                let id = socket_id_of(&handle);
                let state = get_state();
                let mut guard = state.lock().unwrap();
                match guard.tcp_streams.get_mut(&id) {
                    Some(stream) => {
                        // The socket MUST be non-blocking before the read.
                        //
                        // A socket that came from `connect` is blocking, so
                        // `read` on a peer that has sent nothing yet parks the
                        // thread FOREVER — and because the VM drives host calls
                        // synchronously, that hangs the whole program, not just
                        // this call. The `WouldBlock` arm below was written for
                        // a non-blocking socket and was simply never reachable.
                        //
                        // `listen` already does this to its listener; doing it
                        // here too means "no bytes yet" is an EMPTY stream,
                        // which is the answer a `stream<u8>` can express, where
                        // a hang is not an answer at all.
                        let _ = stream.set_nonblocking(true);
                        let mut chunk = [0u8; 65536];
                        match stream.read(&mut chunk) {
                            Ok(read) => buffer.extend_from_slice(&chunk[..read]),
                            Err(error)
                                if error.kind() == std::io::ErrorKind::WouldBlock
                                    || error.kind() == std::io::ErrorKind::TimedOut => {}
                            Err(error) => failure = socket_err_from(&error),
                        }
                    }
                    None => failure = socket_err("invalid-state"),
                }
            } else {
                failure = socket_err("invalid-state");
            }

            let (stream_val, stream_id) = ctx.create_stream();
            for byte in &buffer {
                ctx.stream_push(stream_id, Value::I32(*byte as i32));
            }
            if matches!(failure, Value::Null) {
                // Left OPEN, with a producer, for the same reason `listen`'s
                // stream is: closing here made "the peer has not sent anything
                // YET" indistinguishable from "the peer is gone". A guest
                // reading a request off a socket saw clean EOF on an empty
                // buffer and concluded the connection had ended, when in fact
                // it had not started. The producer blocks for the bytes.
                //
                // §receive: "the implementation drops the stream once no more
                // data is available" — so the producer, not this call, is what
                // closes it, and only on a real end-of-stream.
                if let Some(handle) = socket_arg(args) {
                    get_state()
                        .lock()
                        .unwrap()
                        .recv_streams
                        .insert(stream_id, socket_id_of(&handle));
                    ctx.set_stream_producer(stream_id, PRODUCER_MODULE, PRODUCER_RECEIVE);
                }
            } else {
                // A failed receive has nothing more to give.
                ctx.stream_close(stream_id);
            }
            let (future_val, future_id) = ctx.create_future();
            ctx.resolve_future(future_id, failure);
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
                stream_val, future_val,
            ])))
        }),
    );

    register_socket_address_getters_0_3(vm, "tcp-socket");

    // `get-is-listening: func() -> bool` — infallible, so it must answer on a
    // socket that has only been created.
    vm.register_host_fn(
        "wasi:sockets/types",
        "[method]tcp-socket.get-is-listening",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::Bool(matches!(
                socket_arg(args).and_then(|handle| handle_string(&handle, "__state")),
                Some(state) if state == "listening"
            ))
        }),
    );

    // `set-listen-backlog-size: func(value: u64) -> result<_, error-code>`
    // Recorded for `listen` to apply, because POSIX takes the backlog at
    // `listen(2)` and there is no socket option for it.
    vm.register_host_fn(
        "wasi:sockets/types",
        "[method]tcp-socket.set-listen-backlog-size",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            with_socket(args, |handle, _id| {
                let value = method_arg(args, 0).map(|v| v.as_f64()).unwrap_or(0.0);
                if value < 1.0 {
                    return socket_err("invalid-argument");
                }
                set_handle(handle, "__listen_backlog", Value::F64(value));
                Value::Null
            })
        }),
    );

    register_tcp_socket_options_0_3(vm);
}

/// `get-local-address` / `get-remote-address` / `get-address-family`, shared by
/// both socket resources — the wording and behaviour are identical in 0.3.
fn register_socket_address_getters_0_3(vm: &mut VM, resource: &'static str) {
    for (suffix, key) in [
        ("get-local-address", "__local_address"),
        ("get-remote-address", "__remote_address"),
    ] {
        vm.register_host_fn(
            "wasi:sockets/types",
            &format!("[method]{resource}.{suffix}"),
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let Some(handle) = socket_arg(args) else {
                    return socket_err("invalid-state");
                };
                match handle.lock().unwrap().properties.get(key) {
                    Some(Value::Null) | None => socket_err("invalid-state"),
                    Some(value) => value.clone(),
                }
            }),
        );
    }

    // `get-address-family: func() -> ip-address-family` — infallible.
    vm.register_host_fn(
        "wasi:sockets/types",
        &format!("[method]{resource}.get-address-family"),
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let family = socket_arg(args)
                .and_then(|handle| handle_string(&handle, "__address_family"))
                .unwrap_or_else(|| "ipv4".to_string());
            Value::String(Arc::from(family.as_str()))
        }),
    );
}

/// Is this handle's socket IPv6? Decides TTL vs IPv6 unicast hops.
fn handle_is_ipv6(args: &[Value]) -> bool {
    socket_arg(args)
        .and_then(|handle| handle_string(&handle, "__address_family"))
        .map(|family| family == "ipv6")
        .unwrap_or(false)
}

/// `duration` is u64 NANOSECONDS in `wasi:clocks/types`.
fn duration_arg(args: &[Value], index: usize) -> Option<Duration> {
    let nanos = method_arg(args, index)?.as_f64();
    if nanos < 0.0 {
        return None;
    }
    Some(Duration::from_nanos(nanos as u64))
}

fn duration_value(duration: Duration) -> Value {
    Value::F64(duration.as_nanos() as f64)
}

fn io_result(result: std::io::Result<()>) -> Value {
    match result {
        Ok(()) => Value::Null,
        Err(error) => socket_err_from(&error),
    }
}

/// The `tcp-socket` options: keep-alive, hop limit and buffer sizes.
///
/// `std::net` exposes none of these beyond TTL, which is why `socket2` is a
/// dependency — the alternative was answering `not-supported` to twelve of the
/// interface's functions.
fn register_tcp_socket_options_0_3(vm: &mut VM) {
    vm.register_host_fn(
        "wasi:sockets/types",
        "[method]tcp-socket.get-keep-alive-enabled",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            with_socket(args, |_handle, id| {
                with_socket2_view(id, |socket| match socket.keepalive() {
                    Ok(enabled) => Value::Bool(enabled),
                    Err(error) => socket_err_from(&error),
                })
            })
        }),
    );

    vm.register_host_fn(
        "wasi:sockets/types",
        "[method]tcp-socket.set-keep-alive-enabled",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            with_socket(args, |_handle, id| {
                let enabled = method_arg(args, 0)
                    .map(|value| value.as_bool())
                    .unwrap_or(false);
                with_socket2_view(id, |socket| io_result(socket.set_keepalive(enabled)))
            })
        }),
    );

    vm.register_host_fn(
        "wasi:sockets/types",
        "[method]tcp-socket.get-keep-alive-idle-time",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            with_socket(args, |_handle, id| {
                with_socket2_view(id, |socket| match socket.keepalive_time() {
                    Ok(duration) => duration_value(duration),
                    Err(error) => socket_err_from(&error),
                })
            })
        }),
    );

    vm.register_host_fn(
        "wasi:sockets/types",
        "[method]tcp-socket.set-keep-alive-idle-time",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            with_socket(args, |_handle, id| {
                let Some(duration) = duration_arg(args, 0) else {
                    return socket_err("invalid-argument");
                };
                with_socket2_view(id, |socket| {
                    io_result(
                        socket.set_tcp_keepalive(&socket2::TcpKeepalive::new().with_time(duration)),
                    )
                })
            })
        }),
    );

    vm.register_host_fn(
        "wasi:sockets/types",
        "[method]tcp-socket.get-keep-alive-interval",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            with_socket(args, |_handle, id| {
                with_socket2_view(id, |socket| match socket.keepalive_interval() {
                    Ok(duration) => duration_value(duration),
                    Err(error) => socket_err_from(&error),
                })
            })
        }),
    );

    vm.register_host_fn(
        "wasi:sockets/types",
        "[method]tcp-socket.set-keep-alive-interval",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            with_socket(args, |_handle, id| {
                let Some(duration) = duration_arg(args, 0) else {
                    return socket_err("invalid-argument");
                };
                with_socket2_view(id, |socket| {
                    io_result(
                        socket.set_tcp_keepalive(
                            &socket2::TcpKeepalive::new().with_interval(duration),
                        ),
                    )
                })
            })
        }),
    );

    vm.register_host_fn(
        "wasi:sockets/types",
        "[method]tcp-socket.get-keep-alive-count",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            with_socket(args, |_handle, id| {
                with_socket2_view(id, |socket| match socket.keepalive_retries() {
                    Ok(count) => Value::F64(count as f64),
                    Err(error) => socket_err_from(&error),
                })
            })
        }),
    );

    vm.register_host_fn(
        "wasi:sockets/types",
        "[method]tcp-socket.set-keep-alive-count",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            with_socket(args, |_handle, id| {
                let count = method_arg(args, 0)
                    .map(|value| value.as_f64())
                    .unwrap_or(0.0);
                if count < 1.0 {
                    return socket_err("invalid-argument");
                }
                with_socket2_view(id, |socket| {
                    io_result(socket.set_tcp_keepalive(
                        &socket2::TcpKeepalive::new().with_retries(count as u32),
                    ))
                })
            })
        }),
    );

    // `hop-limit` is IPv4's TTL and IPv6's unicast hop limit — one interface
    // function over two socket options, chosen by the socket's family.
    vm.register_host_fn(
        "wasi:sockets/types",
        "[method]tcp-socket.get-hop-limit",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let ipv6 = handle_is_ipv6(args);
            with_socket(args, |_handle, id| {
                with_socket2_view(id, |socket| {
                    let hops = if ipv6 {
                        socket.unicast_hops_v6()
                    } else {
                        socket.ttl()
                    };
                    match hops {
                        Ok(value) => Value::F64(value as f64),
                        Err(error) => socket_err_from(&error),
                    }
                })
            })
        }),
    );

    vm.register_host_fn(
        "wasi:sockets/types",
        "[method]tcp-socket.set-hop-limit",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let ipv6 = handle_is_ipv6(args);
            with_socket(args, |_handle, id| {
                let value = method_arg(args, 0).map(|v| v.as_f64()).unwrap_or(0.0);
                // "A value of 0 is not allowed" — types.wit, `set-hop-limit`.
                if !(1.0..=255.0).contains(&value) {
                    return socket_err("invalid-argument");
                }
                with_socket2_view(id, |socket| {
                    io_result(if ipv6 {
                        socket.set_unicast_hops_v6(value as u32)
                    } else {
                        socket.set_ttl(value as u32)
                    })
                })
            })
        }),
    );

    register_buffer_size_options_0_3(vm, "tcp-socket");
}

/// `get/set-receive-buffer-size` and `get/set-send-buffer-size` — identical on
/// both socket resources.
fn register_buffer_size_options_0_3(vm: &mut VM, resource: &'static str) {
    vm.register_host_fn(
        "wasi:sockets/types",
        &format!("[method]{resource}.get-receive-buffer-size"),
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            with_socket(args, |_handle, id| {
                with_socket2_view(id, |socket| match socket.recv_buffer_size() {
                    Ok(size) => Value::F64(size as f64),
                    Err(error) => socket_err_from(&error),
                })
            })
        }),
    );

    vm.register_host_fn(
        "wasi:sockets/types",
        &format!("[method]{resource}.set-receive-buffer-size"),
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            with_socket(args, |_handle, id| {
                let size = method_arg(args, 0).map(|v| v.as_f64()).unwrap_or(0.0);
                if size < 1.0 {
                    return socket_err("invalid-argument");
                }
                with_socket2_view(id, |socket| {
                    io_result(socket.set_recv_buffer_size(size as usize))
                })
            })
        }),
    );

    vm.register_host_fn(
        "wasi:sockets/types",
        &format!("[method]{resource}.get-send-buffer-size"),
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            with_socket(args, |_handle, id| {
                with_socket2_view(id, |socket| match socket.send_buffer_size() {
                    Ok(size) => Value::F64(size as f64),
                    Err(error) => socket_err_from(&error),
                })
            })
        }),
    );

    vm.register_host_fn(
        "wasi:sockets/types",
        &format!("[method]{resource}.set-send-buffer-size"),
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            with_socket(args, |_handle, id| {
                let size = method_arg(args, 0).map(|v| v.as_f64()).unwrap_or(0.0);
                if size < 1.0 {
                    return socket_err("invalid-argument");
                }
                with_socket2_view(id, |socket| {
                    io_result(socket.set_send_buffer_size(size as usize))
                })
            })
        }),
    );
}

/// The `udp-socket` resource.
///
/// Deliberately NOT unified with `tcp-socket`: 0.3 gives UDP plain
/// `list<u8>` payloads rather than streams, and `receive` answers the peer
/// address alongside the datagram because a UDP socket hears from many peers.
fn register_udp_socket_0_3(vm: &mut VM) {
    vm.register_host_fn(
        "wasi:sockets/types",
        "[static]udp-socket.create",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let family = family_arg(args, 0);
            let socket = match socket2::Socket::new(
                socket2_domain(&family),
                socket2::Type::DGRAM,
                Some(socket2::Protocol::UDP),
            ) {
                Ok(socket) => socket,
                Err(error) => return socket_err_from(&error),
            };
            let handle = new_socket_handle("udp-socket", &family);
            if let Value::Object(object) = &handle {
                let id = socket_id_of(object);
                get_state().lock().unwrap().raw_sockets.insert(id, socket);
            }
            handle
        }),
    );

    // `bind` commits the datagram socket immediately — unlike TCP there is no
    // later listen/connect decision, so it moves straight into `udp_sockets`.
    vm.register_host_fn(
        "wasi:sockets/types",
        "[method]udp-socket.bind",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            with_socket(args, |handle, id| {
                let Some(address) = method_arg(args, 0).and_then(resolve_socket_addr) else {
                    return socket_err("invalid-argument");
                };
                let state = get_state();
                let mut guard = state.lock().unwrap();
                let Some(socket) = guard.raw_sockets.remove(&id) else {
                    return socket_err("invalid-state");
                };
                if let Err(error) = socket.bind(&address.into()) {
                    guard.raw_sockets.insert(id, socket);
                    return socket_err_from(&error);
                }
                let bound = socket
                    .local_addr()
                    .ok()
                    .and_then(|addr| addr.as_socket())
                    .map(socket_addr_to_value);
                guard.udp_sockets.insert(id, UdpSocket::from(socket));
                drop(guard);

                set_handle(handle, "__state", Value::String(Arc::from("bound")));
                if let Some(bound) = bound {
                    set_handle(handle, "__local_address", bound);
                }
                Value::Null
            })
        }),
    );

    // `connect` on a datagram socket only fixes the default peer.
    vm.register_host_fn(
        "wasi:sockets/types",
        "[method]udp-socket.connect",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            with_socket(args, |handle, id| {
                let Some(address) = method_arg(args, 0).and_then(resolve_socket_addr) else {
                    return socket_err("invalid-argument");
                };
                let state = get_state();
                let guard = state.lock().unwrap();
                let outcome = match guard.udp_sockets.get(&id) {
                    Some(socket) => io_result(socket.connect(address)),
                    None => match guard.raw_sockets.get(&id) {
                        Some(socket) => io_result(socket.connect(&address.into())),
                        None => socket_err("invalid-state"),
                    },
                };
                drop(guard);
                if matches!(outcome, Value::Null) {
                    set_handle(handle, "__remote_address", socket_addr_to_value(address));
                }
                outcome
            })
        }),
    );

    vm.register_host_fn(
        "wasi:sockets/types",
        "[method]udp-socket.disconnect",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            with_socket(args, |handle, id| {
                let state = get_state();
                let guard = state.lock().unwrap();
                let Some(socket) = guard.udp_sockets.get(&id) else {
                    return socket_err("invalid-state");
                };
                let outcome = io_result(udp_disassociate(socket));
                drop(guard);
                if matches!(outcome, Value::Null) {
                    set_handle(handle, "__remote_address", Value::Null);
                }
                outcome
            })
        }),
    );

    // `send: async func(data: list<u8>, remote-address: option<ip-socket-address>)`
    vm.register_host_fn(
        "wasi:sockets/types",
        "[method]udp-socket.send",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            with_socket(args, |_handle, id| {
                let bytes = match method_arg(args, 0) {
                    Some(value) => bytes_from_value(value),
                    None => Vec::new(),
                };
                let remote = method_arg(args, 1).and_then(resolve_socket_addr);
                let state = get_state();
                let guard = state.lock().unwrap();
                let Some(socket) = guard.udp_sockets.get(&id) else {
                    return socket_err("invalid-state");
                };
                let sent = match remote {
                    Some(address) => socket.send_to(&bytes, address),
                    None => socket.send(&bytes),
                };
                match sent {
                    Ok(_) => Value::Null,
                    Err(error) => socket_err_from(&error),
                }
            })
        }),
    );

    // `receive: async func() -> result<tuple<list<u8>, ip-socket-address>, error-code>`
    vm.register_host_fn(
        "wasi:sockets/types",
        "[method]udp-socket.receive",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            with_socket(args, |_handle, id| {
                let state = get_state();
                let guard = state.lock().unwrap();
                let Some(socket) = guard.udp_sockets.get(&id) else {
                    return socket_err("invalid-state");
                };
                // Non-blocking for the same reason as `tcp-socket.receive`: a
                // blocking `recv_from` with no datagram pending parks the VM's
                // only thread forever.
                //
                // "Nothing yet" is reported as `timeout`. 0.2 had a
                // `would-block` error-code and 0.3.1 DELETED it — `receive` is
                // an `async func` there, so not-ready is the future not having
                // resolved, not an error. This VM drives host calls
                // synchronously and has no such future, and `timeout` is the
                // only DECLARED variant that means "no data within the time I
                // was willing to wait". Inventing `would-block` back would be
                // the same offence as the verbs this migration deleted.
                let _ = socket.set_nonblocking(true);
                let mut chunk = [0u8; 65536];
                match socket.recv_from(&mut chunk) {
                    Ok((read, peer)) => {
                        let payload: Vec<Value> = chunk[..read]
                            .iter()
                            .map(|byte| Value::I32(*byte as i32))
                            .collect();
                        Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
                            Value::Object(vybe_runtime::heap::alloc(Object::new_array(payload))),
                            socket_addr_to_value(peer),
                        ])))
                    }
                    Err(error) if error.kind() == std::io::ErrorKind::WouldBlock => {
                        socket_err("timeout")
                    }
                    Err(error) => socket_err_from(&error),
                }
            })
        }),
    );

    register_socket_address_getters_0_3(vm, "udp-socket");

    // UDP's hop limit is spelled `unicast-hop-limit`; same two socket options.
    vm.register_host_fn(
        "wasi:sockets/types",
        "[method]udp-socket.get-unicast-hop-limit",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let ipv6 = handle_is_ipv6(args);
            with_socket(args, |_handle, id| {
                with_socket2_view_udp(id, |socket| {
                    let hops = if ipv6 {
                        socket.unicast_hops_v6()
                    } else {
                        socket.ttl()
                    };
                    match hops {
                        Ok(value) => Value::F64(value as f64),
                        Err(error) => socket_err_from(&error),
                    }
                })
            })
        }),
    );

    vm.register_host_fn(
        "wasi:sockets/types",
        "[method]udp-socket.set-unicast-hop-limit",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let ipv6 = handle_is_ipv6(args);
            with_socket(args, |_handle, id| {
                let value = method_arg(args, 0).map(|v| v.as_f64()).unwrap_or(0.0);
                if !(1.0..=255.0).contains(&value) {
                    return socket_err("invalid-argument");
                }
                with_socket2_view_udp(id, |socket| {
                    io_result(if ipv6 {
                        socket.set_unicast_hops_v6(value as u32)
                    } else {
                        socket.set_ttl(value as u32)
                    })
                })
            })
        }),
    );

    register_buffer_size_options_udp_0_3(vm);
}

/// Dissolve a datagram socket's association with its default peer.
///
/// POSIX spells this `connect(2)` to a `sockaddr` whose family is `AF_UNSPEC`
/// — NOT to the unspecified address. Connecting a UDP socket to `0.0.0.0:0`
/// fails with `EADDRNOTAVAIL` on BSD-derived systems, so that shortcut looks
/// right and is not. Neither `std::net::UdpSocket` nor `socket2` can express an
/// `AF_UNSPEC` address, which is why this reaches for the raw call.
#[cfg(unix)]
fn udp_disassociate(socket: &UdpSocket) -> std::io::Result<()> {
    use std::os::fd::AsRawFd;

    let mut address: libc::sockaddr = unsafe { std::mem::zeroed() };
    address.sa_family = libc::AF_UNSPEC as libc::sa_family_t;
    let result = unsafe {
        libc::connect(
            socket.as_raw_fd(),
            &address as *const libc::sockaddr,
            std::mem::size_of::<libc::sockaddr>() as libc::socklen_t,
        )
    };
    if result == 0 {
        return Ok(());
    }
    let error = std::io::Error::last_os_error();
    // Several BSDs report `EAFNOSUPPORT` from this call while still having
    // performed the disassociation, which is long-standing documented
    // behaviour rather than a failure.
    if error.raw_os_error() == Some(libc::EAFNOSUPPORT) {
        return Ok(());
    }
    Err(error)
}

#[cfg(not(unix))]
fn udp_disassociate(socket: &UdpSocket) -> std::io::Result<()> {
    socket.connect(SocketAddr::from((Ipv4Addr::UNSPECIFIED, 0)))
}

/// Datagram sockets live in `udp_sockets` once bound and in `raw_sockets`
/// before that; both are the same OS socket seen through `socket2`.
fn with_socket2_view_udp(id: u64, body: impl FnOnce(&socket2::Socket) -> Value) -> Value {
    let state = get_state();
    let guard = state.lock().unwrap();
    if let Some(socket) = guard.raw_sockets.get(&id) {
        return body(socket);
    }
    if let Some(socket) = guard.udp_sockets.get(&id) {
        return body(&socket2::SockRef::from(socket));
    }
    socket_err("invalid-state")
}

/// The buffer-size options against the UDP socket maps.
fn register_buffer_size_options_udp_0_3(vm: &mut VM) {
    for (suffix, apply) in [
        (
            "get-receive-buffer-size",
            (|socket: &socket2::Socket, _size: Option<usize>| match socket.recv_buffer_size() {
                Ok(size) => Value::F64(size as f64),
                Err(error) => socket_err_from(&error),
            }) as fn(&socket2::Socket, Option<usize>) -> Value,
        ),
        ("get-send-buffer-size", |socket, _size| {
            match socket.send_buffer_size() {
                Ok(size) => Value::F64(size as f64),
                Err(error) => socket_err_from(&error),
            }
        }),
        ("set-receive-buffer-size", |socket, size| {
            io_result(socket.set_recv_buffer_size(size.unwrap_or(0)))
        }),
        ("set-send-buffer-size", |socket, size| {
            io_result(socket.set_send_buffer_size(size.unwrap_or(0)))
        }),
    ] {
        let is_setter = suffix.starts_with("set-");
        vm.register_host_fn(
            "wasi:sockets/types",
            &format!("[method]udp-socket.{suffix}"),
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                with_socket(args, |_handle, id| {
                    let size = if is_setter {
                        let raw = method_arg(args, 0).map(|v| v.as_f64()).unwrap_or(0.0);
                        if raw < 1.0 {
                            return socket_err("invalid-argument");
                        }
                        Some(raw as usize)
                    } else {
                        None
                    };
                    with_socket2_view_udp(id, |socket| apply(socket, size))
                })
            }),
        );
    }
}

pub fn register_wasi_sockets_0_3(vm: &mut VM) {
    register_tcp_socket_0_3(vm);
    register_udp_socket_0_3(vm);

    // `resolve-addresses: async func(name: string) -> result<list<ip-address>,
    // error-code>`
    //
    // `ip-name-lookup` is the one sibling interface 0.3.1 keeps, but the SHAPE
    // changed: 0.2 answered a `resolve-address-stream` RESOURCE that callers
    // drained through `resolve-next-address`, and this tree modelled that as an
    // object carrying `__addresses`. 0.3.1 answers the list directly, so that
    // is what this returns. Python's `gethostbyname` already indexed the result
    // with `ARRAY_GET 0` — it was written against the list and was wrong only
    // because the host handed back the resource.
    //
    // The `network` argument 0.2 took first is gone with `instance-network`,
    // so the name is the sole argument; `args.last()` reads it either way.
    vm.register_host_fn(
        "wasi:sockets/ip-name-lookup",
        "resolve-addresses",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let name = args.last().map(|v| format!("{}", v)).unwrap_or_default();
            let addresses = resolve_name_socket_addrs(&name)
                .into_iter()
                .map(socket_addr_to_value)
                .collect();
            value_array_inner(addresses)
        }),
    );
}













fn resolve_name_socket_addrs(host: &str) -> Vec<SocketAddr> {
    match std::net::ToSocketAddrs::to_socket_addrs(&format!("{}:0", host)) {
        Ok(addrs) => addrs.collect(),
        Err(_) => Vec::new(),
    }
}

pub fn make_pollable(target: Arc<Mutex<Object>>) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("Pollable")));
    obj.properties
        .insert("__target".into(), Value::Object(target));
    Value::Object(vybe_runtime::heap::alloc(obj))
}








fn socket_arg(args: &[Value]) -> Option<Arc<Mutex<Object>>> {
    match args.first() {
        Some(Value::Object(obj)) => Some(obj.clone()),
        // A `borrow<tcp-socket>` self parameter lowers as an i32 handle, and a
        // guest that read its socket off `listen`'s `stream<tcp-socket>` has
        // exactly that and nothing else. Before this it reached here as an
        // integer, matched no arm, and every method on an ACCEPTED connection
        // answered `invalid-state` — the socket was fine, the argument was
        // simply in the shape the canonical ABI says it should be.
        Some(Value::I32(rep)) => socket_reps().lock().unwrap().get(&(*rep as u64)).cloned(),
        Some(Value::F64(rep)) => socket_reps().lock().unwrap().get(&(*rep as u64)).cloned(),
        _ => None,
    }
}

fn method_arg<'a>(args: &'a [Value], index: usize) -> Option<&'a Value> {
    // A self parameter is whatever [`socket_arg`] can resolve — an object, or
    // the i32 handle the canonical ABI actually passes. Testing for `Object`
    // alone made an accepted socket's `bind(addr)` read the SOCKET as its
    // address argument.
    let offset = usize::from(socket_arg(args).is_some());
    args.get(offset + index)
}




fn as_object_value(value: &Value) -> Option<Arc<Mutex<Object>>> {
    match value {
        Value::Object(obj) => Some(obj.clone()),
        _ => None,
    }
}

fn stream_kind(stream: &Arc<Mutex<Object>>) -> Option<String> {
    match stream.lock().unwrap().properties.get("__type") {
        Some(Value::String(kind)) => Some(kind.to_string()),
        _ => None,
    }
}

pub fn stream_socket_id(stream: &Arc<Mutex<Object>>) -> Option<u64> {
    match stream.lock().unwrap().properties.get("__socket_id") {
        Some(Value::F64(id)) => Some(*id as u64),
        Some(Value::I64(id)) => Some(*id as u64),
        Some(Value::I32(id)) => Some(*id as u64),
        _ => None,
    }
}

pub fn value_len_arg(value: &Value) -> usize {
    let len = value.as_i64();
    len.clamp(0, 64 * 1024) as usize
}

pub fn byte_array(bytes: &[u8]) -> Value {
    value_array(bytes.iter().map(|byte| Value::I32(*byte as i32)).collect())
}

pub fn bytes_from_value(value: &Value) -> Vec<u8> {
    // One owner: ECMA (§23.2) answers what a value's bytes are. The
    // local copy this replaced CLAMPED to 0..255 where node truncated —
    // same Array, different bytes per platform — and neither handled a
    // TypedArray, so a Python `bytes` written to a socket sent NOTHING.
    vybe_platform_ecma::typedarray::bytes_from_value(value)
}

pub fn read_stream_bytes(stream: &Arc<Mutex<Object>>, len: usize, blocking: bool) -> Value {
    let Some(socket_id) = stream_socket_id(stream) else {
        return Value::Null;
    };
    if len == 0 {
        return value_array(vec![]);
    }

    let state = get_state();
    let mut guard = state.lock().unwrap();
    let Some(tcp_stream) = guard.tcp_streams.get_mut(&socket_id) else {
        return Value::Null;
    };

    let mut buf = vec![0u8; len];
    if !blocking {
        let _ = tcp_stream.set_nonblocking(true);
    }
    let result = match tcp_stream.read(&mut buf) {
        Ok(count) => byte_array(&buf[..count]),
        Err(err) if !blocking && err.kind() == std::io::ErrorKind::WouldBlock => {
            value_array(vec![])
        }
        Err(err) => {
            eprintln!("tcp-socket.receive error: {}", err);
            Value::Null
        }
    };
    if !blocking {
        let _ = tcp_stream.set_nonblocking(false);
    }
    result
}

pub fn write_stream_bytes(stream: &Arc<Mutex<Object>>, contents: &[u8], flush: bool) -> Value {
    let Some(socket_id) = stream_socket_id(stream) else {
        return Value::Null;
    };
    let state = get_state();
    let mut guard = state.lock().unwrap();
    let Some(tcp_stream) = guard.tcp_streams.get_mut(&socket_id) else {
        return Value::Null;
    };

    match tcp_stream.write_all(contents) {
        Ok(()) => {
            if flush {
                let _ = tcp_stream.flush();
            }
            Value::Null
        }
        Err(err) => {
            eprintln!("tcp-socket.send error: {}", err);
            Value::Null
        }
    }
}

pub fn flush_stream(stream: &Arc<Mutex<Object>>) -> Value {
    let Some(socket_id) = stream_socket_id(stream) else {
        return Value::Null;
    };
    let state = get_state();
    let mut guard = state.lock().unwrap();
    let Some(tcp_stream) = guard.tcp_streams.get_mut(&socket_id) else {
        return Value::Null;
    };
    match tcp_stream.flush() {
        Ok(()) => Value::Null,
        Err(err) => {
            eprintln!("tcp-socket.send flush error: {}", err);
            Value::Null
        }
    }
}

pub fn splice_streams(
    src: &Arc<Mutex<Object>>,
    dst: &Arc<Mutex<Object>>,
    len: usize,
    blocking: bool,
) -> Value {
    let bytes = read_stream_bytes(src, len, blocking);
    let Value::Object(array) = &bytes else {
        return Value::Null;
    };
    let contents = bytes_from_value(&Value::Object(array.clone()));
    if write_stream_bytes(dst, &contents, blocking).type_tag() == "null" {
        Value::I64(contents.len() as i64)
    } else {
        Value::Null
    }
}

pub fn value_array_elements(value: &Value) -> Option<Vec<Value>> {
    match value {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            if let vybe_runtime::value::ObjectKind::Array(elements) = &obj.kind {
                Some(elements.clone())
            } else {
                None
            }
        }
        _ => None,
    }
}

pub fn array_len(array: &Arc<Mutex<Object>>) -> usize {
    let array = array.lock().unwrap();
    if let vybe_runtime::value::ObjectKind::Array(elements) = &array.kind {
        elements.len()
    } else {
        0
    }
}


pub fn pollable_ready(pollable: &Arc<Mutex<Object>>) -> bool {
    let target = {
        let pollable = pollable.lock().unwrap();
        pollable
            .properties
            .get("__target")
            .cloned()
            .unwrap_or(Value::Null)
    };
    let Some(target) = as_object_value(&target) else {
        return stream_or_socket_ready(pollable);
    };
    stream_or_socket_ready(&target)
}

pub fn block_until_ready(pollable: &Arc<Mutex<Object>>) {
    let start = Instant::now();
    while !pollable_ready(pollable) && start.elapsed() < Duration::from_secs(1) {
        thread::sleep(Duration::from_millis(1));
    }
}

fn stream_or_socket_ready(resource: &Arc<Mutex<Object>>) -> bool {
    match stream_kind(resource).as_deref() {
        Some("InputStream") => input_stream_ready(resource),
        Some("OutputStream") => stream_socket_id(resource).is_some(),
        Some("TcpSocket") => tcp_socket_ready(resource),
        Some("Pollable") => pollable_ready(resource),
        Some("TimerPollable") => {
            use std::sync::OnceLock;
            static START: OnceLock<Instant> = OnceLock::new();
            let start = START.get_or_init(Instant::now);
            let ready_at = resource
                .lock()
                .unwrap()
                .properties
                .get("__ready_at_ns")
                .map(|v| v.as_f64() as u128)
                .unwrap_or(0);
            start.elapsed().as_nanos() >= ready_at
        }
        _ => false,
    }
}

fn input_stream_ready(stream: &Arc<Mutex<Object>>) -> bool {
    let Some(socket_id) = stream_socket_id(stream) else {
        return false;
    };
    let state = get_state();
    let mut guard = state.lock().unwrap();
    let Some(tcp_stream) = guard.tcp_streams.get_mut(&socket_id) else {
        return false;
    };

    let _ = tcp_stream.set_nonblocking(true);
    let mut buf = [0u8; 1];
    let ready = match tcp_stream.peek(&mut buf) {
        Ok(_) => true,
        Err(err) if err.kind() == std::io::ErrorKind::WouldBlock => false,
        Err(_) => true,
    };
    let _ = tcp_stream.set_nonblocking(false);
    ready
}

fn tcp_socket_ready(socket: &Arc<Mutex<Object>>) -> bool {
    let listener_id = {
        let socket = socket.lock().unwrap();
        socket.properties.get("__listener_id").cloned()
    };
    let Some(Value::F64(listener_id)) = listener_id else {
        return stream_socket_id(socket).is_some();
    };
    let listener_id = listener_id as u64;

    let state = get_state();
    let mut guard = state.lock().unwrap();
    if guard
        .pending_accepts
        .get(&listener_id)
        .map(|pending| !pending.is_empty())
        .unwrap_or(false)
    {
        return true;
    }

    let accepted = if let Some(listener) = guard.tcp_listeners.get(&listener_id) {
        match listener.accept() {
            Ok((stream, _)) => Some(stream),
            Err(err) if err.kind() == std::io::ErrorKind::WouldBlock => None,
            Err(_) => None,
        }
    } else {
        None
    };

    if let Some(stream) = accepted {
        guard
            .pending_accepts
            .entry(listener_id)
            .or_default()
            .push(stream);
        return true;
    }

    false
}


fn parse_ip_socket_address(value: &Value) -> Option<(String, u16, String)> {
    match value {
        Value::String(s) => parse_socket_addr_str(s),
        Value::Object(obj) => {
            let props = &obj.lock().unwrap().properties;
            if let Some(Value::Object(inner)) = props.get("ipv4").or_else(|| props.get("ipv6")) {
                let family = if props.contains_key("ipv6") {
                    "ipv6"
                } else {
                    "ipv4"
                };
                return parse_ip_socket_record(&inner.lock().unwrap().properties, family);
            }
            parse_ip_socket_record(props, "ipv4")
        }
        _ => None,
    }
}

fn parse_ip_socket_record(
    props: &vybe_runtime::Properties,
    default_family: &str,
) -> Option<(String, u16, String)> {
    let port = props.get("port")?.as_f64() as u16;
    if let Some(Value::String(address)) = props.get("address") {
        return Some((address.to_string(), port, default_family.to_string()));
    }
    if let Some(Value::Object(address)) = props.get("address") {
        let address_obj = address.lock().unwrap();
        if let vybe_runtime::value::ObjectKind::Array(elems) = &address_obj.kind {
            let family = if elems.len() == 4 { "ipv4" } else { "ipv6" };
            let host = if elems.len() == 4 {
                format!(
                    "{}.{}.{}.{}",
                    elems[0].as_i32(),
                    elems[1].as_i32(),
                    elems[2].as_i32(),
                    elems[3].as_i32()
                )
            } else {
                elems
                    .iter()
                    .map(|elem| format!("{:x}", elem.as_i32()))
                    .collect::<Vec<_>>()
                    .join(":")
            };
            return Some((host, port, family.into()));
        }
    }
    None
}

fn parse_socket_addr_str(input: &str) -> Option<(String, u16, String)> {
    if let Ok(addr) = input.parse::<SocketAddr>() {
        let family = match addr.ip() {
            IpAddr::V4(_) => "ipv4",
            IpAddr::V6(_) => "ipv6",
        };
        return Some((addr.ip().to_string(), addr.port(), family.into()));
    }
    if let Some((host, port)) = input.rsplit_once(':') {
        if let Ok(port) = port.parse::<u16>() {
            let family = if host.contains(':') { "ipv6" } else { "ipv4" };
            return Some((host.to_string(), port, family.into()));
        }
    }
    None
}

fn socket_addr_to_value(addr: SocketAddr) -> Value {
    make_ip_socket_address(
        &addr.ip().to_string(),
        addr.port(),
        if addr.is_ipv6() { "ipv6" } else { "ipv4" },
    )
}

fn make_ip_socket_address(host: &str, port: u16, family: &str) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("family".into(), Value::String(Arc::from(family)));
    obj.properties
        .insert("port".into(), Value::F64(port as f64));
    obj.properties
        .insert("address".into(), ip_address_value(host, family));
    Value::Object(vybe_runtime::heap::alloc(obj))
}

fn ip_address_value(host: &str, family: &str) -> Value {
    let values = if family == "ipv6" {
        host.parse::<Ipv6Addr>()
            .map(|addr| {
                addr.segments()
                    .into_iter()
                    .map(|segment| Value::I32(segment as i32))
                    .collect()
            })
            .unwrap_or_else(|_| vec![Value::I32(0); 8])
    } else {
        host.parse::<Ipv4Addr>()
            .map(|addr| {
                addr.octets()
                    .into_iter()
                    .map(|octet| Value::I32(octet as i32))
                    .collect()
            })
            .unwrap_or_else(|_| vec![Value::I32(127), Value::I32(0), Value::I32(0), Value::I32(1)])
    };
    value_array_inner(values)
}

pub fn value_array(elements: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(elements)))
}

fn value_array_inner(elements: Vec<Value>) -> Value {
    value_array(elements)
}

// ── Phase 3: spec-correct [method] prefix forms ───────────────────────────────
//
// Each existing flat registration (e.g. `wasi:sockets/tcp`, `start-bind`) is
// mirrored as `wasi:sockets/tcp`, `[method]tcp-socket.start-bind`. Both forms
// stay registered so existing callers are not broken.

