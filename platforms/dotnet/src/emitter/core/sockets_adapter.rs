//! .NET sockets adapter — bytecode-only `.NET → WASI sockets` translation.
//!
//! Every `System.Net.Sockets.*` and `System.Net.Dns.*` method that the
//! .NET wrapper exposes lowers through one of the `Common` emits in
//! this file. The emit produces bytecode that calls the standardized
//! `wasi:sockets/{types,ip-name-lookup}.*` and `node:os.*` host
//! primitives — no `dotnet:*` host module involved.
//!
//! WASI 0.3.1 collapsed `tcp`, `tcp-create-socket`, `udp`,
//! `udp-create-socket`, `network` and `instance-network` into the single
//! `wasi:sockets/types` interface, and the two-phase `start-*`/`finish-*`
//! pairs into one call each. Two .NET verbs have NO 0.3.1 spelling and are
//! marked `GAP:` at their emit — see each doc comment.
//!
//! Architectural rule: the host exposes only spec-aligned namespaces
//! (`ecma:*`, `wasi:*`, `wasm:*`, `web:*`, `node:*`). Anything
//! .NET-shaped lives here at compile time as an emitter adapter. The
//! .NET surface still looks .NET-shaped to the source code (`New
//! TcpClient(host, port)`, `listener.AcceptTcpClient()`, etc.); only
//! the underlying bytecode is standardized.
//!
//! Each emit assumes the user-supplied args are already on the stack
//! in source order (per the `MethodBody::Common` calling convention
//! shared with the rest of `compiler_common::dispatch`).

use std::sync::Arc;
use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use super::object_fields::field_slot;

/// Emit `CONST <idx>` for a literal value — `Chunk` doesn't expose
/// this directly the way the compiler's `emit_const` helper does,
/// so we inline the two-step (add_constant + emit_op_u16) pattern.
fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        _ => panic!("push_const: no WASM-compliant encoding for {:?}", val),
    }
}

/// Build a `"host:port"` IP-socket-address string on the stack.
/// Inputs from local slots; output is a single `String` value pushed.
/// Stack: `[]` → `[String("host:port")]`. Uses `Op::DYN_ADD` with dynamic
/// string-concat semantics (string + number → string per JS).
fn emit_host_port_string(chunk: &mut Chunk, host_slot: u16, port_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, host_slot, line);
    push_const(chunk, Value::String(Arc::from(":")), line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, port_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
}

fn struct_set_drop(chunk: &mut Chunk, field: &str, line: u32) {
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(field),
        ValueSource::Stack,
        line,
    );
}

fn struct_get(chunk: &mut Chunk, field: &str, line: u32) {
    class_slots::emit_class_get(
        chunk,
        ObjSource::Stack,
        &field_slot(field),
        Dest::Stack,
        line,
    );
}

// ─── Dns ─────────────────────────────────────────────────────────────────

/// `Dns.GetHostAddresses(host)` — resolves `host` to an array of IP
/// address strings.
///
/// Stack: `[host]` → `[array<string>]`
///
/// Lowers to `wasi:sockets/ip-name-lookup.resolve-addresses(host)`, which in
/// 0.3.1 answers `list<ip-address>` DIRECTLY. 0.2 answered a
/// `resolve-address-stream` resource that had to be drained, and this emit
/// used to unwrap its `__addresses` field; there is no resource to unwrap
/// any more, so the call result is already the array.
pub fn emit_dns_get_host_addresses(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let resolve_idx =
        chunks[current].add_import("wasi:sockets/ip-name-lookup", "resolve-addresses");
    chunks[current].emit_call(resolve_idx, 1, line);
}

/// `Dns.GetHostEntry(host)` — resolves `host` and wraps the result in
/// the .NET `IPHostEntry` shape: `{ HostName, AddressList, Aliases }`.
pub fn emit_dns_get_host_entry(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let resolve_idx =
        chunks[current].add_import("wasi:sockets/ip-name-lookup", "resolve-addresses");
    let lower_idx = chunks[current].add_import("ecma:string", "toLowerCase");
    let chunk = &mut chunks[current];
    let host_slot = chunk.alloc_scratch(3);
    let addresses_slot = host_slot + 1;
    let obj_slot = host_slot + 2;

    chunk.emit_op_u16(Op::LOCAL_SET, host_slot, line);

    // 0.3.1 hands back `list<ip-address>`; nothing to unwrap.
    chunk.emit_op_u16(Op::LOCAL_GET, host_slot, line);
    chunk.emit_call(resolve_idx, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, addresses_slot, line);

    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("IPHostEntry")), line);
    struct_set_drop(chunk, "__type", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, host_slot, line);
    chunk.emit_call(lower_idx, 1, line);
    struct_set_drop(chunk, "HostName", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, addresses_slot, line);
    struct_set_drop(chunk, "AddressList", line);

    core_wasm::dup(chunk, line);
    chunk.emit_array_new_fixed(0, 0, line);
    struct_set_drop(chunk, "Aliases", line);

    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

/// `Dns.GetHostName()` — returns the local machine's hostname.
/// Lowers to `node:os.hostname()`.
///
/// Stack: `[]` → `[string]`
pub fn emit_dns_get_host_name(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let idx = chunks[current].add_import("node:os", "hostname");
    chunks[current].emit_call(idx, 0, line);
}

pub fn emit_ip_address_parse(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let text_slot = chunk.alloc_scratch(2);
    let obj_slot = text_slot + 1;

    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);

    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("IPAddress")), line);
    struct_set_drop(chunk, "__type", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    struct_set_drop(chunk, "__value", line);

    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("InterNetwork")), line);
    struct_set_drop(chunk, "AddressFamily", line);

    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

pub fn emit_ip_address_to_string(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let typeof_idx = chunks[current].add_import("ecma:value", "typeof");
    let undefined_idx = chunks[current].add_import("wasm:js-undefined", "test");
    let stringify_idx = chunks[current].add_import("ecma:string", "String");
    let chunk = &mut chunks[current];
    let receiver_slot = chunk.alloc_scratch(2);
    let value_slot = receiver_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, receiver_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunk.emit_call(typeof_idx, 1, line);
    push_const(chunk, Value::String(Arc::from("string")), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    struct_get(chunk, "__value", line);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_call(undefined_idx, 1, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunk.emit_call(stringify_idx, 1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

// ─── TcpClient ───────────────────────────────────────────────────────────

/// `New TcpClient(host, port)` — synchronous connect.
///
/// Stack: `[host, port]` → `[tcp_socket]`
///
/// Composition (per WASI 0.3.1 wasi-sockets):
///   1. `[static]tcp-socket.create(ipv4)` → socket
///   2. `[method]tcp-socket.connect(socket, "host:port")` — one call; 0.2's
///      `start-connect`/`finish-connect` pair is gone, and so is the
///      `network` handle that used to sit between the socket and the address
///   3. Stamp `__type=TcpClient` on the socket so runtime dispatch
///      finds the .NET adapter TypeDef
pub fn emit_tcp_client_new(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    // Stack at entry: [host, port] (user args)
    // Stash to scratch locals so we can re-push in any order.
    let host_slot = chunks[current].alloc_scratch(2);
    let port_slot = host_slot + 1;
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, port_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, host_slot, line);

    // 1. Create socket
    let create_idx = chunks[current].add_import("wasi:sockets/types", "[static]tcp-socket.create");
    push_const(&mut chunks[current], Value::String(Arc::from("ipv4")), line);
    chunks[current].emit_call(create_idx, 1, line);
    // Stack: [socket]

    // Stash the socket for return + as `connect`'s receiver.
    let sock_slot = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, sock_slot, line);

    // 2. Push args for connect: (socket, "host:port"). 0.3.1 dropped the
    //    `network` argument along with the `instance-network` interface.
    chunk.emit_op_u16(Op::LOCAL_GET, sock_slot, line);
    emit_host_port_string(chunk, host_slot, port_slot, line);
    // Stack: [socket, "host:port"]

    let connect_idx = chunks[current].add_import("wasi:sockets/types", "[method]tcp-socket.connect");
    chunks[current].emit_call(connect_idx, 2, line);
    chunks[current].emit_op(Op::DROP, line); // discard connect result

    // 3. Re-push socket as the value of `New TcpClient`. The runtime
    //    `__stamp_type` stamp added by the dotnet ctor flow tags it
    //    with `__type=TcpClient`.
    chunks[current].emit_op_u16(Op::LOCAL_GET, sock_slot, line);
}

/// `tcpClient.GetStream()` — return the endpoint reads and writes go through.
///
/// Stack: `[client]` → `[tcp_socket]`
///
/// 0.2 took this pair from `tcp.finish-connect`, which returned
/// `[input_stream, output_stream]`. 0.3.1 has no `finish-connect` at all:
/// `connect` returns nothing, and the byte streams are reached through
/// `[method]tcp-socket.receive` (`tuple<stream<u8>, future<...>>`) and
/// `[method]tcp-socket.send` (`stream<u8>` in). The socket IS the endpoint,
/// so `GetStream()` hands the socket straight back and a future
/// `NetworkStream.Read`/`Write` lowers to `receive`/`send` on it.
pub fn emit_tcp_client_get_stream(_chunks: &mut Vec<Chunk>, _current: usize, _line: u32) {
    // socket already on the stack; it is its own read/write endpoint.
}

/// `tcpClient.Close()` — release the socket.
///
/// Stack: `[client]` → `[null]` (void return)
///
/// GAP: 0.2's `tcp.shutdown` is DELETED in 0.3.1 and has no replacement
/// function — the Component Model closes a socket by dropping its resource
/// handle (`canon resource.drop`). Vybe's socket handles are plain objects
/// carrying `__socket_id`, not typed resource handles, so `resource.drop`
/// rejects them ("handle is not a resource handle") and there is no
/// spec-legal lowering to emit. Dropping the value is therefore all this
/// can do, and the OS socket stays in the host's registry until teardown.
/// Closing it properly needs socket handles to become real resource
/// handles — host/VM work, which is gated on the user's approval.
pub fn emit_tcp_client_close(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

// ─── TcpListener ─────────────────────────────────────────────────────────

/// `New TcpListener(port)` — bind + start listening on `0.0.0.0:port`.
///
/// Stack: `[port]` → `[tcp_listener]`
///
/// Composition (0.3.1 — five calls become three):
///   1. `[static]tcp-socket.create(ipv4)` → socket
///   2. `[method]tcp-socket.bind(socket, "0.0.0.0:port")`
///   3. `[method]tcp-socket.listen(socket)` → `stream<tcp-socket>`
///
/// `listen` answers a STREAM of inbound sockets — that stream is the whole
/// of 0.3.1's accept mechanism. Nothing in the bytecode can read it (see
/// `emit_tcp_listener_accept`), so it is dropped here and the listener
/// socket is what the constructor yields, matching the .NET shape.
pub fn emit_tcp_listener_new(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let port_slot = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, port_slot, line);

    let create_idx = chunks[current].add_import("wasi:sockets/types", "[static]tcp-socket.create");
    push_const(&mut chunks[current], Value::String(Arc::from("ipv4")), line);
    chunks[current].emit_call(create_idx, 1, line);

    let sock_slot = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, sock_slot, line);

    // bind(socket, "0.0.0.0:port") — one call, no `network` handle.
    chunk.emit_op_u16(Op::LOCAL_GET, sock_slot, line);
    push_const(chunk, Value::String(Arc::from("0.0.0.0")), line);
    push_const(chunk, Value::String(Arc::from(":")), line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, port_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    let bind_idx = chunks[current].add_import("wasi:sockets/types", "[method]tcp-socket.bind");
    chunks[current].emit_call(bind_idx, 2, line);
    chunks[current].emit_op(Op::DROP, line);

    // listen(socket) -> stream<tcp-socket>; the stream is unreadable from
    // bytecode, so it is dropped. Calling it still flips the socket into the
    // listening state, which is what `Pending()` and the bound port report on.
    chunks[current].emit_op_u16(Op::LOCAL_GET, sock_slot, line);
    let listen_idx = chunks[current].add_import("wasi:sockets/types", "[method]tcp-socket.listen");
    chunks[current].emit_call(listen_idx, 1, line);
    chunks[current].emit_op(Op::DROP, line);

    // Return the listener socket
    chunks[current].emit_op_u16(Op::LOCAL_GET, sock_slot, line);
}

/// `listener.Start()` — no-op in our wasi-sockets impl (start-listen
/// is part of construction). Returns the listener for chaining.
///
/// Stack: `[listener]` → `[listener]`
pub fn emit_tcp_listener_start(_chunks: &mut Vec<Chunk>, _current: usize, _line: u32) {
    // listener already on stack; pass through.
}

/// `listener.Stop()` — release the listening socket.
///
/// Stack: `[listener]` → `[null]`
///
/// GAP: same as `emit_tcp_client_close` — 0.3.1 deleted `tcp.shutdown` and
/// closes by dropping the resource handle, which Vybe's object-shaped socket
/// handles cannot express.
pub fn emit_tcp_listener_stop(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `listener.AcceptTcpClient()` — return the next connected socket.
///
/// Stack: `[listener]` → `[null]`
///
/// GAP: 0.2's `tcp.accept` is DELETED. In 0.3.1 the ONLY way to take an
/// inbound connection is to read the `stream<tcp-socket>` that
/// `[method]tcp-socket.listen` returned — accept is not a function any more,
/// it is a stream read.
///
/// That read cannot be emitted: `canon stream.read` copies BYTES into linear
/// memory, so it handles `stream<u8>` and nothing else, and
/// `EventLoop::stream_pop` (which does move whole items) has no bytecode
/// opcode reaching it. This is the SAME blocker as
/// `wasi:filesystem`'s `read-directory: stream<directory-entry>` — one
/// architectural item, not two.
///
/// `null` is the pre-existing contract for "no client pending": 0.2's
/// `accept` already returned null in non-blocking mode and the caller
/// indexed it with `ARRAY_GET`, which tolerates null. So this narrows a
/// capability rather than inventing a new silent failure — but it IS a
/// narrowing, and the blocking form stays unavailable until a stream of
/// non-`u8` elements can be read.
pub fn emit_tcp_listener_accept(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `listener.Pending()` — true if `AcceptTcpClient()` would not block.
/// Synchronous query; we approximate by checking the listener's
/// listening state (0.3.1 spells it `get-is-listening`).
///
/// Stack: `[listener]` → `[bool]`
pub fn emit_tcp_listener_pending(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let idx =
        chunks[current].add_import("wasi:sockets/types", "[method]tcp-socket.get-is-listening");
    chunks[current].emit_call(idx, 1, line);
}

// ─── UdpClient ───────────────────────────────────────────────────────────

/// `New UdpClient(port)` — bind a UDP socket to `0.0.0.0:port`.
///
/// Stack: `[port]` → `[udp_socket]`
pub fn emit_udp_client_new(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let port_slot = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, port_slot, line);

    let create_idx = chunks[current].add_import("wasi:sockets/types", "[static]udp-socket.create");
    push_const(&mut chunks[current], Value::String(Arc::from("ipv4")), line);
    chunks[current].emit_call(create_idx, 1, line);

    let sock_slot = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, sock_slot, line);

    // bind(socket, "0.0.0.0:port") — 0.2's start-bind/finish-bind pair.
    chunk.emit_op_u16(Op::LOCAL_GET, sock_slot, line);
    push_const(chunk, Value::String(Arc::from("0.0.0.0")), line);
    push_const(chunk, Value::String(Arc::from(":")), line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, port_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    let bind_idx = chunks[current].add_import("wasi:sockets/types", "[method]udp-socket.bind");
    chunks[current].emit_call(bind_idx, 2, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, sock_slot, line);
}

/// `udp.Send(data, length, host, port)` — send a datagram.
///
/// Stack: `[client, data, length, host, port]` → `[null]`
///
/// 0.2 routed this through `udp.stream`, which returned an
/// `outgoing-datagram-stream` resource you then called `send` on — and the
/// five-argument call this emitted never matched that shape in the first
/// place. 0.3.1 deletes the datagram-stream resources entirely:
/// `[method]udp-socket.send(data: list<u8>, remote-address)` is one call on
/// the socket, which is both honest and simpler.
///
/// .NET's `length` argument is consumed and not forwarded — `send` takes the
/// whole list — exactly as before.
pub fn emit_udp_send(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    // Stack at entry: [client, data, length, host, port] — pop in reverse.
    let sock_slot = chunks[current].alloc_scratch(5);
    let data_slot = sock_slot + 1;
    let len_slot = sock_slot + 2;
    let host_slot = sock_slot + 3;
    let port_slot = sock_slot + 4;
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, port_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, host_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, data_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, sock_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, sock_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, data_slot, line);
    emit_host_port_string(chunk, host_slot, port_slot, line);
    let idx = chunks[current].add_import("wasi:sockets/types", "[method]udp-socket.send");
    chunks[current].emit_call(idx, 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `udp.Receive()` — receive a datagram, returns bytes.
///
/// Stack: `[client]` → `[bytes]`
///
/// `[method]udp-socket.receive()` answers
/// `result<tuple<list<u8>, ip-socket-address>, error-code>`; .NET's
/// `Receive` wants the payload, so element 0 of that tuple.
pub fn emit_udp_receive(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let idx = chunks[current].add_import("wasi:sockets/types", "[method]udp-socket.receive");
    chunks[current].emit_call(idx, 1, line);
    push_const(&mut chunks[current], Value::I32(0), line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
}

/// `udp.Close()` — close the UDP socket.
///
/// Stack: `[client]` → `[null]`
pub fn emit_udp_close(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}
