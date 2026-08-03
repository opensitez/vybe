//! .NET sockets adapter — bytecode-only `.NET → WASI sockets` translation.
//!
//! Every `System.Net.Sockets.*` and `System.Net.Dns.*` method that the
//! .NET wrapper exposes lowers through one of the `Common` emits in
//! this file. The emit produces bytecode that calls the standardized
//! `wasi:sockets/{tcp,udp,ip-name-lookup,instance-network}.*` and
//! `node:os.*` host primitives — no `dotnet:*` host module involved.
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
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};
use vybe_compiler::primitives::instructions::core_wasm;

/// Emit `CONST <idx>` for a literal value — `Chunk` doesn't expose
/// this directly the way the compiler's `emit_const` helper does,
/// so we inline the two-step (add_constant + emit_op_u16) pattern.
fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        _ => panic!("push_const: no WASM-compliant encoding for {:?}", val) }
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
    let key = chunk.add_constant(Value::String(Arc::from(field)));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
    chunk.emit_op(Op::DROP, line);
}

fn struct_get(chunk: &mut Chunk, field: &str, line: u32) {
    let key = chunk.add_constant(Value::String(Arc::from(field)));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
}

// ─── Dns ─────────────────────────────────────────────────────────────────

/// `Dns.GetHostAddresses(host)` — resolves `host` to an array of IP
/// address strings.
///
/// Stack: `[host]` → `[array<string>]`
///
/// Lowers to `wasi:sockets/ip-name-lookup.resolve-addresses(host)` which
/// returns a `ResolveAddressStream` resource. The stream's
/// `__addresses` field is the already-collected array of IP strings;
/// we just `STRUCT_GET` it to drain.
pub fn emit_dns_get_host_addresses(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let resolve_idx =
        chunks[current].add_import("wasi:sockets/ip-name-lookup", "resolve-addresses");
    chunks[current].emit_op_u16(Op::CALL_IMPORT, resolve_idx, line);
    chunks[current].emit(1, line);
    let addrs_key = chunks[current].add_constant(Value::String(Arc::from("__addresses")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, addrs_key, line);
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

    chunk.emit_op_u16(Op::LOCAL_GET, host_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, resolve_idx, line);
    chunk.emit(1, line);
    struct_get(chunk, "__addresses", line);
    chunk.emit_op_u16(Op::LOCAL_SET, addresses_slot, line);

    chunk.emit_struct_new(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("IPHostEntry")), line);
    struct_set_drop(chunk, "__type", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, host_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, lower_idx, line);
    chunk.emit(1, line);
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
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(0, line);
}

pub fn emit_ip_address_parse(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let text_slot = chunk.alloc_scratch(2);
    let obj_slot = text_slot + 1;

    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);

    chunk.emit_struct_new(0, 0, line);
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
    chunk.emit_op_u16(Op::CALL_IMPORT, typeof_idx, line);
    chunk.emit(1, line);
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
    chunk.emit_op_u16(Op::CALL_IMPORT, undefined_idx, line);
    chunk.emit(1, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, stringify_idx, line);
    chunk.emit(1, line);
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
/// Composition (per WASI 0.2.11 wasi-sockets):
///   1. `tcp-create-socket(ipv4)` → socket
///   2. Build IP address record `{ "ipv4": [host, port] }`
///   3. `tcp.start-connect(socket, network=null, addr)` — our impl is
///      synchronous and lenient (uses last arg as remote addr)
///   4. Stamp `__type=TcpClient` on the socket so runtime dispatch
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
    let create_idx =
        chunks[current].add_import("wasi:sockets/tcp-create-socket", "create-tcp-socket");
    push_const(&mut chunks[current], Value::String(Arc::from("ipv4")), line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, create_idx, line);
    chunks[current].emit(1, line);
    // Stack: [socket]

    // Stash the socket for return + for start-connect arg 0.
    let sock_slot = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, sock_slot, line);

    // 2. Push args for start-connect: (socket, network=null, "host:port")
    chunk.emit_op_u16(Op::LOCAL_GET, sock_slot, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    emit_host_port_string(chunk, host_slot, port_slot, line);
    // Stack: [socket, null, "host:port"]

    let connect_idx = chunks[current].add_import("wasi:sockets/tcp", "start-connect");
    chunks[current].emit_op_u16(Op::CALL_IMPORT, connect_idx, line);
    chunks[current].emit(3, line);
    chunks[current].emit_op(Op::DROP, line); // discard connect result

    // 3. Re-push socket as the value of `New TcpClient`. The runtime
    //    `__stamp_type` stamp added by the dotnet ctor flow tags it
    //    with `__type=TcpClient`.
    chunks[current].emit_op_u16(Op::LOCAL_GET, sock_slot, line);
}

/// `tcpClient.GetStream()` — return the (input, output) stream pair.
///
/// Stack: `[client]` → `[stream_pair_array]`
///
/// `wasi:sockets/tcp.finish-connect(socket)` returns
/// `[input_stream, output_stream]` as a 2-element array — exactly what
/// .NET callers feed into `New StreamReader(stream)` /
/// `New StreamWriter(stream)`.
pub fn emit_tcp_client_get_stream(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let idx = chunks[current].add_import("wasi:sockets/tcp", "finish-connect");
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(1, line);
}

/// `tcpClient.Close()` — shut down the socket.
///
/// Stack: `[client]` → `[null]` (void return)
pub fn emit_tcp_client_close(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let idx = chunks[current].add_import("wasi:sockets/tcp", "shutdown");
    push_const(&mut chunks[current], Value::String(Arc::from("both")), line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

// ─── TcpListener ─────────────────────────────────────────────────────────

/// `New TcpListener(port)` — bind + start listening on `0.0.0.0:port`.
///
/// Stack: `[port]` → `[tcp_listener]`
///
/// Composition:
///   1. `tcp-create-socket(ipv4)` → socket
///   2. `tcp.start-bind(socket, network=null, addr={0.0.0.0, port})`
///   3. `tcp.finish-bind(socket)`
///   4. `tcp.start-listen(socket)`
///   5. `tcp.finish-listen(socket)`
pub fn emit_tcp_listener_new(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let port_slot = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, port_slot, line);

    let create_idx =
        chunks[current].add_import("wasi:sockets/tcp-create-socket", "create-tcp-socket");
    push_const(&mut chunks[current], Value::String(Arc::from("ipv4")), line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, create_idx, line);
    chunks[current].emit(1, line);

    let sock_slot = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, sock_slot, line);

    // start-bind(socket, network=null, addr="0.0.0.0:port")
    chunk.emit_op_u16(Op::LOCAL_GET, sock_slot, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    push_const(chunk, Value::String(Arc::from("0.0.0.0")), line);
    push_const(chunk, Value::String(Arc::from(":")), line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, port_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    let bind_idx = chunks[current].add_import("wasi:sockets/tcp", "start-bind");
    chunks[current].emit_op_u16(Op::CALL_IMPORT, bind_idx, line);
    chunks[current].emit(3, line);
    chunks[current].emit_op(Op::DROP, line);

    // finish-bind — synchronous, just acknowledges
    chunks[current].emit_op_u16(Op::LOCAL_GET, sock_slot, line);
    let fbind_idx = chunks[current].add_import("wasi:sockets/tcp", "finish-bind");
    chunks[current].emit_op_u16(Op::CALL_IMPORT, fbind_idx, line);
    chunks[current].emit(1, line);
    chunks[current].emit_op(Op::DROP, line);

    // start-listen
    chunks[current].emit_op_u16(Op::LOCAL_GET, sock_slot, line);
    let listen_idx = chunks[current].add_import("wasi:sockets/tcp", "start-listen");
    chunks[current].emit_op_u16(Op::CALL_IMPORT, listen_idx, line);
    chunks[current].emit(1, line);
    chunks[current].emit_op(Op::DROP, line);

    // finish-listen
    chunks[current].emit_op_u16(Op::LOCAL_GET, sock_slot, line);
    let flisten_idx = chunks[current].add_import("wasi:sockets/tcp", "finish-listen");
    chunks[current].emit_op_u16(Op::CALL_IMPORT, flisten_idx, line);
    chunks[current].emit(1, line);
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

/// `listener.Stop()` — shut down the listening socket.
///
/// Stack: `[listener]` → `[null]`
pub fn emit_tcp_listener_stop(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let idx = chunks[current].add_import("wasi:sockets/tcp", "shutdown");
    push_const(&mut chunks[current], Value::String(Arc::from("both")), line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `listener.AcceptTcpClient()` — block until a client connects, return
/// the connected socket.
///
/// Stack: `[listener]` → `[tcp_socket | null]`
///
/// `wasi:sockets/tcp.accept(listener)` returns `[client_socket,
/// input_stream, output_stream]` array, or `null` if no pending client
/// (non-blocking mode). For .NET semantics we want just the socket;
/// extract index 0.
pub fn emit_tcp_listener_accept(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let accept_idx = chunks[current].add_import("wasi:sockets/tcp", "accept");
    chunks[current].emit_op_u16(Op::CALL_IMPORT, accept_idx, line);
    chunks[current].emit(1, line);
    // Result on stack: [array_of_3] | null
    // Index 0 is the client socket. Use ARRAY_GET (handles null gracefully).
    push_const(&mut chunks[current], Value::I32(0), line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
}

/// `listener.Pending()` — true if `AcceptTcpClient()` would not block.
/// Synchronous query; we approximate by checking the listener's
/// `is-listening` state.
///
/// Stack: `[listener]` → `[bool]`
pub fn emit_tcp_listener_pending(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let idx = chunks[current].add_import("wasi:sockets/tcp", "is-listening");
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(1, line);
}

// ─── UdpClient ───────────────────────────────────────────────────────────

/// `New UdpClient(port)` — bind a UDP socket to `0.0.0.0:port`.
///
/// Stack: `[port]` → `[udp_socket]`
pub fn emit_udp_client_new(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let port_slot = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, port_slot, line);

    let create_idx =
        chunks[current].add_import("wasi:sockets/udp-create-socket", "create-udp-socket");
    push_const(&mut chunks[current], Value::String(Arc::from("ipv4")), line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, create_idx, line);
    chunks[current].emit(1, line);

    let sock_slot = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, sock_slot, line);

    // start-bind(socket, network=null, addr="0.0.0.0:port")
    chunk.emit_op_u16(Op::LOCAL_GET, sock_slot, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    push_const(chunk, Value::String(Arc::from("0.0.0.0")), line);
    push_const(chunk, Value::String(Arc::from(":")), line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, port_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    let bind_idx = chunks[current].add_import("wasi:sockets/udp", "start-bind");
    chunks[current].emit_op_u16(Op::CALL_IMPORT, bind_idx, line);
    chunks[current].emit(3, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, sock_slot, line);
    let fbind_idx = chunks[current].add_import("wasi:sockets/udp", "finish-bind");
    chunks[current].emit_op_u16(Op::CALL_IMPORT, fbind_idx, line);
    chunks[current].emit(1, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, sock_slot, line);
}

/// `udp.Send(data, length, host, port)` — send a datagram.
///
/// Stack: `[client, data, length, host, port]` → `[null]`
///
/// Lowers to `wasi:sockets/udp.stream(socket).send(...)` shape. Vybe's
/// impl exposes the send through the stream resource returned by
/// `udp.stream`; we wrap the call into a single emit for caller
/// simplicity.
pub fn emit_udp_send(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let idx = chunks[current].add_import("wasi:sockets/udp", "stream");
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(5, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `udp.Receive()` — receive a datagram, returns bytes.
///
/// Stack: `[client]` → `[bytes]`
pub fn emit_udp_receive(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let idx = chunks[current].add_import("wasi:sockets/udp", "stream");
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(1, line);
}

/// `udp.Close()` — close the UDP socket.
///
/// Stack: `[client]` → `[null]`
pub fn emit_udp_close(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}
