//! Python `socket` — the module-level functions that need no socket object.
//!
//! Three different backings, picked by what the operation actually is:
//!
//! * `inet_aton` / `inet_ntoa` are pure arithmetic over a dotted quad. No host
//!   call at all — they are string↔bytes conversions, not networking.
//! * `getservbyname` is an `/etc/services` lookup. WASI has no such interface
//!   and neither does node; the well-known ports are a fixed IANA table, so it
//!   is a compile-time table here.
//! * `gethostname` is host identity rather than a socket operation, so it comes
//!   from `node:os.hostname` — `wasi:sockets` has no equivalent.
//! * `gethostbyname` / `getaddrinfo` are real resolution and go to
//!   `wasi:sockets/ip-name-lookup.resolve-addresses`.
//!
//! Nothing here registers a host function.
//!
//! Arguments arrive pre-pushed on the stack, left to right — the `emit_common`
//! convention shared with `url_adapter` / `os_path_adapter`.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use vybe_compiler::primitives::tuples;

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    // Register on the CURRENT chunk — an index taken from chunks[0] resolves
    // to the wrong host fn when the code runs inside a function chunk.
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}

fn lget(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
}

/// Pop `argc` values (deepest first) into fresh scratch slots; `base + i` is
/// the i-th argument. Extra arguments are simply stashed and ignored.
fn stash_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(argc as u16);
    for offset in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    base
}

// ─── the socket OBJECT ───────────────────────────────────────────────────────
//
// `wasi:sockets/types@0.3.1`, emitted directly. This used to be
// `VybeSocketImpl`, a python class in `SOCKET_PRELUDE`, and moving it here is
// not tidying: a prelude is compiled as ORDINARY USER PYTHON, so it inherits
// every walker defect. `_addr_tuple` built its host string with
// `".".join(pieces)` and that call resolved to UNDEFINED the moment a program
// said `import threading`, because type-directed dispatch preferred the
// prelude's own `Thread.join` over the string built-in. Bytecode cannot be
// hijacked that way.
//
// The receiver is the host socket handle itself — there is no wrapper object.
// `__type` on that handle ("tcp-socket" / "udp-socket") is what the tcp/udp
// branches below read, so the adapter never has to guess which resource it is
// holding.

/// The interface every 0.3.1 socket method lives on. 0.3.1 collapsed `tcp`,
/// `udp`, `tcp-create-socket`, `udp-create-socket` and `instance-network` into
/// this one.
const SOCK: &str = "wasi:sockets/types";

fn get_prop(chunks: &mut [Chunk], current: usize, obj: u16, key: &str, line: u32) {
    lget(chunks, current, obj, line);
    chunks[current].emit_string_const(key, line);
    call_import(chunks, current, "ecma:object", "get", 2, line);
}

fn set_prop(chunks: &mut [Chunk], current: usize, obj: u16, key: &str, line: u32) {
    // Stack on entry: [value]. Object and key go under it via a scratch hop.
    let value = chunks[current].alloc_scratch(1);
    lset(chunks, current, value, line);
    lget(chunks, current, obj, line);
    chunks[current].emit_string_const(key, line);
    lget(chunks, current, value, line);
    call_import(chunks, current, "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// Leaves `true` when the handle in `obj` is a `udp-socket`.
fn emit_is_udp(chunks: &mut [Chunk], current: usize, obj: u16, line: u32) {
    get_prop(chunks, current, obj, "__type", line);
    chunks[current].emit_string_const("udp-socket", line);
    call_import(chunks, current, "wasm:js-string", "equals", 2, line);
}

/// A python address `(host, port)` → the `"host:port"` text the host parses as
/// an `ip-socket-address`.
///
/// A LOOP, not `":".join(...)` — see the module note above.
fn emit_addr_text(chunks: &mut [Chunk], current: usize, addr: u16, line: u32) {
    lget(chunks, current, addr, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    chunks[current].emit_string_const(":", line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    lget(chunks, current, addr, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
}

/// A WIT `ip-socket-address` record → python's `(host, port)`.
///
/// `address` is a list of octets for ipv4 and a string for ipv6, which is why
/// the shape is tested rather than assumed. A record that carries no address
/// at all — what `get-local-address` answers on an unbound socket, as an
/// `error-code` rather than a null — degrades to `("0.0.0.0", 0)`, the same
/// answer CPython gives for an unbound socket.
fn emit_addr_tuple(chunks: &mut [Chunk], current: usize, rec: u16, line: u32) {
    let parts = chunks[current].alloc_scratch(1);
    let host = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);

    get_prop(chunks, current, rec, "address", line);
    lset(chunks, current, parts, line);

    lget(chunks, current, parts, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    chunks[current].emit_if_value(line);
    {
        // ipv4: dotted quad, built by hand.
        chunks[current].emit_string_const("", line);
        lset(chunks, current, host, line);
        chunks[current].emit_i32_const(0, line);
        lset(chunks, current, idx, line);

        let done = chunks[current].emit_block(line);
        let (loop_id, _) = chunks[current].emit_loop_s(line);
        lget(chunks, current, idx, line);
        lget(chunks, current, parts, line);
        chunks[current].emit_op(Op::ARRAY_LENGTH, line);
        chunks[current].emit_op(Op::I32_GE_S, line);
        chunks[current].emit_br_if(1, line);

        lget(chunks, current, host, line);
        lget(chunks, current, idx, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op(Op::I32_GT_S, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_string_const(".", line);
        chunks[current].emit_else(line);
        chunks[current].emit_string_const("", line);
        chunks[current].emit_end(line);
        call_import(chunks, current, "wasm:js-string", "concat", 2, line);
        lget(chunks, current, parts, line);
        lget(chunks, current, idx, line);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        call_import(chunks, current, "ecma:string", "String", 1, line);
        call_import(chunks, current, "wasm:js-string", "concat", 2, line);
        lset(chunks, current, host, line);

        lget(chunks, current, idx, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        lset(chunks, current, idx, line);
        chunks[current].emit_br(0, line);
        chunks[current].emit_end(line);
        chunks[current].patch_loop(loop_id);
        chunks[current].emit_end(line);
        chunks[current].patch_block(done);

        lget(chunks, current, host, line);
    }
    chunks[current].emit_else(line);
    {
        // ipv6 gives the address as text already; a missing address gives the
        // unspecified one rather than a crash inside the caller's format.
        lget(chunks, current, parts, line);
        call_import(chunks, current, "ecma:value", "typeof", 1, line);
        chunks[current].emit_string_const("string", line);
        call_import(chunks, current, "wasm:js-string", "equals", 2, line);
        chunks[current].emit_if_value(line);
        lget(chunks, current, parts, line);
        chunks[current].emit_else(line);
        chunks[current].emit_string_const("0.0.0.0", line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_end(line);

    get_prop(chunks, current, rec, "port", line);
    tuples::emit_tuple(chunks, current, 2, line);
}

/// `socket.socket(family, kind, proto)` → a real `tcp-socket`/`udp-socket`.
/// Stack: `[family, kind, proto?] -> [handle]`.
pub fn emit_sock_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    // `SOCK_DGRAM` is 2. Defaulting to STREAM matches CPython's own default.
    if argc >= 2 {
        lget(chunks, current, base + 1, line);
        chunks[current].emit_f64_const(2.0, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    } else {
        chunks[current].emit_bool_const(false, line);
    }
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("ipv4", line);
    call_import(chunks, current, SOCK, "[static]udp-socket.create", 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("ipv4", line);
    call_import(chunks, current, SOCK, "[static]tcp-socket.create", 1, line);
    chunks[current].emit_end(line);
}

/// `sock.bind((host, port))`. Stack: `[sock, addr] -> [None]`.
pub fn emit_sock_bind(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let (sock, addr) = (base, base + 1);
    emit_is_udp(chunks, current, sock, line);
    chunks[current].emit_if_value(line);
    lget(chunks, current, sock, line);
    emit_addr_text(chunks, current, addr, line);
    call_import(chunks, current, SOCK, "[method]udp-socket.bind", 2, line);
    chunks[current].emit_else(line);
    lget(chunks, current, sock, line);
    emit_addr_text(chunks, current, addr, line);
    call_import(chunks, current, SOCK, "[method]tcp-socket.bind", 2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `sock.listen(backlog)`.
///
/// The returned `stream<tcp-socket>` is STASHED on the handle rather than
/// returned: python's `listen()` answers None and it is `accept()` that needs
/// the stream. 0.3.1 has no `accept` — the stream IS the accept queue.
pub fn emit_sock_listen(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let sock = base;
    if argc >= 2 {
        lget(chunks, current, sock, line);
        lget(chunks, current, base + 1, line);
        call_import(
            chunks,
            current,
            SOCK,
            "[method]tcp-socket.set-listen-backlog-size",
            2,
            line,
        );
        chunks[current].emit_op(Op::DROP, line);
    }
    lget(chunks, current, sock, line);
    call_import(chunks, current, SOCK, "[method]tcp-socket.listen", 1, line);
    set_prop(chunks, current, sock, "__listener", line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `sock.accept()` → `(conn, addr)`.
///
/// One element read off the listen stream. `common:stream.read_handle` does
/// the `canon stream.read` + `canon resource.rep`; the rep is the socket id,
/// which every method here accepts as its receiver.
pub fn emit_sock_accept(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let sock = base;
    let conn = chunks[current].alloc_scratch(1);

    // A bare `accept()` on a socket nobody called `listen()` on still has to
    // work: python allows it after `bind`, and the stream is what makes the
    // connection reachable at all.
    get_prop(chunks, current, sock, "__listener", line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    lget(chunks, current, sock, line);
    call_import(chunks, current, SOCK, "[method]tcp-socket.listen", 1, line);
    set_prop(chunks, current, sock, "__listener", line);
    chunks[current].emit_end(line);

    get_prop(chunks, current, sock, "__listener", line);
    vybe_compiler::primitives::io::emit_read_stream_handle(&mut chunks[current], line);
    lset(chunks, current, conn, line);

    lget(chunks, current, conn, line);
    lget(chunks, current, conn, line);
    call_import(
        chunks,
        current,
        SOCK,
        "[method]tcp-socket.get-remote-address",
        1,
        line,
    );
    let rec = chunks[current].alloc_scratch(1);
    lset(chunks, current, rec, line);
    emit_addr_tuple(chunks, current, rec, line);
    tuples::emit_tuple(chunks, current, 2, line);
}

/// `sock.connect((host, port))`.
pub fn emit_sock_connect(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let (sock, addr) = (base, base + 1);
    emit_is_udp(chunks, current, sock, line);
    chunks[current].emit_if_value(line);
    lget(chunks, current, sock, line);
    emit_addr_text(chunks, current, addr, line);
    call_import(chunks, current, SOCK, "[method]udp-socket.connect", 2, line);
    chunks[current].emit_else(line);
    lget(chunks, current, sock, line);
    emit_addr_text(chunks, current, addr, line);
    call_import(chunks, current, SOCK, "[method]tcp-socket.connect", 2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `sock.send(data)` / `sendall(data)`.
///
/// 0.3.1's `send` takes `data: stream<u8>`, not a `list<u8>` — bytes are no
/// longer handed to a sink, a stream is produced and passed.
/// `common:stream.from_bytes` mints it.
pub fn emit_sock_send(chunks: &mut [Chunk], current: usize, argc: u8, line: u32, count: bool) {
    let base = stash_args(chunks, current, argc, line);
    let (sock, data) = (base, base + 1);
    emit_is_udp(chunks, current, sock, line);
    chunks[current].emit_if_value(line);
    {
        // UDP still takes the bytes directly: a datagram is one message, so
        // there is nothing for a stream to express.
        lget(chunks, current, sock, line);
        lget(chunks, current, data, line);
        call_import(chunks, current, SOCK, "[method]udp-socket.send", 2, line);
    }
    chunks[current].emit_else(line);
    {
        lget(chunks, current, sock, line);
        lget(chunks, current, data, line);
        vybe_compiler::primitives::io::emit_bytes_to_stream(&mut chunks[current], line);
        call_import(chunks, current, SOCK, "[method]tcp-socket.send", 2, line);
    }
    chunks[current].emit_end(line);
    chunks[current].emit_op(Op::DROP, line);

    if count {
        lget(chunks, current, data, line);
        chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
}

/// `sock.recv(bufsize)`.
///
/// `receive()` is called ONCE per socket and its `stream<u8>` cached on the
/// handle — the WIT says "may be called at most once". Each `recv` is then one
/// bounded `canon stream.read`, which is what makes it return as soon as
/// anything arrives instead of at disconnect.
pub fn emit_sock_recv(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let sock = base;

    emit_is_udp(chunks, current, sock, line);
    chunks[current].emit_if_value(line);
    {
        lget(chunks, current, sock, line);
        call_import(chunks, current, SOCK, "[method]udp-socket.receive", 1, line);
    }
    chunks[current].emit_else(line);
    {
        get_prop(chunks, current, sock, "__rx", line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if(line);
        lget(chunks, current, sock, line);
        call_import(chunks, current, SOCK, "[method]tcp-socket.receive", 1, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        set_prop(chunks, current, sock, "__rx", line);
        chunks[current].emit_end(line);

        get_prop(chunks, current, sock, "__rx", line);
        if argc >= 2 {
            lget(chunks, current, base + 1, line);
        } else {
            chunks[current].emit_i32_const(1024, line);
        }
        vybe_compiler::primitives::io::emit_read_stream_chunk(&mut chunks[current], line);
    }
    chunks[current].emit_end(line);
}

/// `sock.getsockname()` / `getpeername()`.
pub fn emit_sock_addr(chunks: &mut [Chunk], current: usize, argc: u8, line: u32, local: bool) {
    let base = stash_args(chunks, current, argc, line);
    let rec = chunks[current].alloc_scratch(1);
    lget(chunks, current, base, line);
    call_import(
        chunks,
        current,
        SOCK,
        if local {
            "[method]tcp-socket.get-local-address"
        } else {
            "[method]tcp-socket.get-remote-address"
        },
        1,
        line,
    );
    lset(chunks, current, rec, line);
    emit_addr_tuple(chunks, current, rec, line);
}

/// `sock.close()` / `shutdown()`.
///
/// 0.3.1 deleted `shutdown`, and closing a resource is `canon resource.drop`,
/// not an interface call. Releasing the cached receive stream is what actually
/// ends the conversation here.
pub fn emit_sock_close(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    set_prop(chunks, current, base, "__rx", line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    set_prop(chunks, current, base, "__listener", line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// Bookkeeping that never reaches the network: `settimeout`, `setblocking`,
/// `setsockopt` and their readers. Stored ON the handle so a `dup()` and the
/// original agree, which is what CPython does (they share a descriptor).
pub fn emit_sock_setopt(chunks: &mut [Chunk], current: usize, argc: u8, key: &str, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    if argc >= 2 {
        lget(chunks, current, base + argc as u16 - 1, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
    set_prop(chunks, current, base, key, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_sock_getopt(chunks: &mut [Chunk], current: usize, argc: u8, key: &str, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    get_prop(chunks, current, base, key, line);
}

/// `sock.fileno()` — the socket id, which is this component's chosen
/// representation for the resource and the only integer that names it.
pub fn emit_sock_fileno(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    get_prop(chunks, current, base, "__socket_id", line);
}

/// `sock.dup()` / `makefile()` / `__enter__` — all answer the handle itself.
/// A dup shares the descriptor, and `makefile` on this surface reads and
/// writes the same socket.
pub fn emit_sock_self(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    lget(chunks, current, base, line);
}

/// `inet_aton("192.168.1.1")` → the 4 packed bytes, as Python `bytes`
/// (a `Uint8Array`, the shape `bytes` uses everywhere else here).
/// Stack: `[str] -> [bytes]`.
pub fn emit_inet_aton(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
    }
    let base = stash_args(chunks, current, argc, line);

    // parts = s.split(".")
    let parts = chunks[current].alloc_scratch(1);
    lget(chunks, current, base, line);
    chunks[current].emit_string_const(".", line);
    call_import(chunks, current, "ecma:string", "split", 2, line);
    lset(chunks, current, parts, line);

    // Build the 4 octets as numbers, then wrap in a Uint8Array so `len()` is 4
    // and indexing yields ints — the same representation `b"..."` produces.
    for octet in 0..4u16 {
        lget(chunks, current, parts, line);
        chunks[current].emit_i32_const(octet as i32, line);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        // `parseInt`, not `wasm:js-number.toF64` — the split pieces are
        // STRINGS, and `toF64` is a cast that traps on one rather than
        // coercing it.
        chunks[current].emit_f64_const(10.0, line);
        call_import(chunks, current, "ecma:number", "parseInt", 2, line);
    }
    chunks[current].emit_array_new_fixed(0, 4, line);
    // Same constructor `int.to_bytes` / the `b"..."` literal path uses, so
    // `len()` and indexing behave identically to any other `bytes` value.
    call_import(chunks, current, "ecma:uint8array", "new", 1, line);
}

/// `inet_ntoa(b"\xc0\xa8\x01\x01")` → `"192.168.1.1"`.
/// Stack: `[bytes] -> [str]`.
pub fn emit_inet_ntoa(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
    let base = stash_args(chunks, current, argc, line);
    let out = chunks[current].alloc_scratch(1);

    chunks[current].emit_string_const("", line);
    lset(chunks, current, out, line);
    for octet in 0..4i32 {
        lget(chunks, current, out, line);
        if octet > 0 {
            chunks[current].emit_string_const(".", line);
            call_import(chunks, current, "wasm:js-string", "concat", 2, line);
        }
        lget(chunks, current, base, line);
        chunks[current].emit_i32_const(octet, line);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        call_import(chunks, current, "ecma:string", "String", 1, line);
        call_import(chunks, current, "wasm:js-string", "concat", 2, line);
        lset(chunks, current, out, line);
    }
    lget(chunks, current, out, line);
}

/// The IANA well-known ports `getservbyname` is asked for in practice. This is
/// `/etc/services` data, not an interface: WASI has no name-to-port lookup and
/// node exposes none either, so the table is the implementation.
const WELL_KNOWN_PORTS: &[(&str, i32)] = &[
    ("ftp-data", 20),
    ("ftp", 21),
    ("ssh", 22),
    ("telnet", 23),
    ("smtp", 25),
    ("domain", 53),
    ("http", 80),
    ("pop3", 110),
    ("ntp", 123),
    ("imap", 143),
    ("snmp", 161),
    ("ldap", 389),
    ("https", 443),
    ("smtps", 465),
    ("submission", 587),
    ("imaps", 993),
    ("pop3s", 995),
    ("mysql", 3306),
    ("postgresql", 5432),
    ("redis", 6379),
];

/// `getservbyname("http")` → `80`. An unknown name yields `-1` rather than
/// trapping, so a caller can tell "not in the table" from a real port.
/// Stack: `[name, …] -> [number]`.
pub fn emit_getservbyname(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
    }
    let base = stash_args(chunks, current, argc, line);
    let name = chunks[current].alloc_scratch(1);
    lget(chunks, current, base, line);
    call_import(chunks, current, "ecma:string", "toLowerCase", 1, line);
    lset(chunks, current, name, line);

    // A chain of equality tests — the table is fixed at compile time, so this
    // is a constant-folded lookup, not a runtime data structure.
    for (service, port) in WELL_KNOWN_PORTS {
        lget(chunks, current, name, line);
        chunks[current].emit_string_const(service, line);
        call_import(chunks, current, "wasm:js-string", "equals", 2, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_f64_const(*port as f64, line);
        chunks[current].emit_else(line);
    }
    chunks[current].emit_f64_const(-1.0, line);
    for _ in WELL_KNOWN_PORTS {
        chunks[current].emit_end(line);
    }
}

/// `gethostname()` — host identity, so `node:os.hostname`. `wasi:sockets` has
/// no equivalent; the WASI answer would be `wasi:cli/environment`, which does
/// not carry a hostname either.
/// Stack: `[] -> [str]`.
pub fn emit_gethostname(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    call_import(chunks, current, "node:os", "hostname", 0, line);
}

/// Resolve `host` through `wasi:sockets/ip-name-lookup`, leaving the FIRST
/// address as a dotted-quad string in the pushed value.
///
/// `localhost` is answered without a lookup: the resolver may not be reachable
/// in a sandbox, and the loopback answer is fixed.
/// Stack: `[host, …] -> [str]`.
pub fn emit_gethostbyname(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
    }
    let base = stash_args(chunks, current, argc, line);
    let host = base;

    lget(chunks, current, host, line);
    chunks[current].emit_string_const("localhost", line);
    call_import(chunks, current, "wasm:js-string", "equals", 2, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("127.0.0.1", line);
    chunks[current].emit_else(line);
    {
        let list = chunks[current].alloc_scratch(1);
        // 0.3.1: `resolve-addresses: async func(name: string)`. The `network`
        // handle 0.2 took as its first argument is gone with the
        // `instance-network` interface, so the name is the only argument.
        lget(chunks, current, host, line);
        call_import(
            chunks,
            current,
            "wasi:sockets/ip-name-lookup",
            "resolve-addresses",
            1,
            line,
        );
        lset(chunks, current, list, line);

        // No answer → hand back the name unchanged, which is what CPython does
        // for an already-numeric argument.
        lget(chunks, current, list, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if_value(line);
        lget(chunks, current, host, line);
        chunks[current].emit_else(line);
        lget(chunks, current, list, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_end(line);
}

/// `getaddrinfo(host, port)` → `[(family, type, proto, canonname, sockaddr)]`,
/// with one entry for the resolved address. CPython returns a list of 5-tuples
/// whose last element is itself the `(ip, port)` tuple.
/// Stack: `[host, port, …] -> [list]`.
pub fn emit_getaddrinfo(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
    }
    let base = stash_args(chunks, current, argc.max(1), line);
    let ip = chunks[current].alloc_scratch(1);

    // Reuse the resolution above by pushing the host and calling into it.
    lget(chunks, current, base, line);
    emit_gethostbyname(chunks, current, 1, line);
    lset(chunks, current, ip, line);

    // AF_INET, SOCK_STREAM, IPPROTO_TCP, "", (ip, port)
    chunks[current].emit_f64_const(2.0, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_f64_const(6.0, line);
    chunks[current].emit_string_const("", line);
    lget(chunks, current, ip, line);
    if argc >= 2 {
        lget(chunks, current, base + 1, line);
    } else {
        chunks[current].emit_f64_const(0.0, line);
    }
    tuples::emit_tuple(chunks, current, 2, line);
    tuples::emit_tuple(chunks, current, 5, line);
    chunks[current].emit_array_new_fixed(0, 1, line);
}

// ── ipaddress helpers ───────────────────────────────────────────────────────
//
// The `ipaddress` prelude is pure Python except for these: the dotted-quad and
// hextet conversions, which are string/number work the prelude would express
// far more slowly.

/// `_vybe_ip4_parse("192.168.1.1")` → the 32-bit integer.
/// Stack: `[str] -> [number]`.
pub fn emit_ip4_parse(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
    }
    let base = stash_args(chunks, current, argc, line);
    let parts = chunks[current].alloc_scratch(1);
    lget(chunks, current, base, line);
    chunks[current].emit_string_const(".", line);
    call_import(chunks, current, "ecma:string", "split", 2, line);
    lset(chunks, current, parts, line);

    // ((a * 256 + b) * 256 + c) * 256 + d, built left to right.
    for octet in 0..4i32 {
        if octet > 0 {
            chunks[current].emit_f64_const(256.0, line);
            chunks[current].emit_op(Op::F64_MUL, line);
        }
        lget(chunks, current, parts, line);
        chunks[current].emit_i32_const(octet, line);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        chunks[current].emit_f64_const(10.0, line);
        call_import(chunks, current, "ecma:number", "parseInt", 2, line);
        if octet > 0 {
            chunks[current].emit_op(Op::F64_ADD, line);
        }
    }
}

/// Leave the four octets of the 32-bit value in `value` on the stack, most
/// significant first.
///
/// The per-octet reduction is `math::emit_c_fmod` — the shared `a - trunc(a/b)
/// * b`, pure WASM opcodes. WASM has no `f64.rem` (remainder is integer-only:
/// `i32.rem_s/u`, `i64.rem_s/u`), and ECMA spells remainder as the `%`
/// OPERATOR rather than a `Math` function, so there is no `ecma:math` import to
/// reach for. Both operands are non-negative here, so C truncation and floor
/// modulo agree.
fn push_ip4_octets(chunks: &mut [Chunk], current: usize, value: u16, line: u32) {
    for shift in [16777216.0f64, 65536.0, 256.0, 1.0] {
        lget(chunks, current, value, line);
        chunks[current].emit_f64_const(shift, line);
        chunks[current].emit_op(Op::F64_DIV, line);
        chunks[current].emit_op(Op::F64_FLOOR, line);
        chunks[current].emit_f64_const(256.0, line);
        vybe_compiler::primitives::math::emit_c_fmod(&mut chunks[current], line);
    }
}

/// `_vybe_ip4_str(n)` → `"192.168.1.1"`. Stack: `[number] -> [str]`.
pub fn emit_ip4_str(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_f64_const(0.0, line);
    }
    let value = chunks[current].alloc_scratch(1);
    lset(chunks, current, value, line);
    let octets = chunks[current].alloc_scratch(1);
    push_ip4_octets(chunks, current, value, line);
    chunks[current].emit_array_new_fixed(0, 4, line);
    lset(chunks, current, octets, line);

    let out = chunks[current].alloc_scratch(1);
    chunks[current].emit_string_const("", line);
    lset(chunks, current, out, line);
    for index in 0..4i32 {
        lget(chunks, current, out, line);
        if index > 0 {
            chunks[current].emit_string_const(".", line);
            call_import(chunks, current, "wasm:js-string", "concat", 2, line);
        }
        lget(chunks, current, octets, line);
        chunks[current].emit_i32_const(index, line);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        call_import(chunks, current, "ecma:string", "String", 1, line);
        call_import(chunks, current, "wasm:js-string", "concat", 2, line);
        lset(chunks, current, out, line);
    }
    lget(chunks, current, out, line);
}

/// `_vybe_ip4_octets(n)` → `[a, b, c, d]`. Stack: `[number] -> [array]`.
pub fn emit_ip4_octets(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_f64_const(0.0, line);
    }
    let value = chunks[current].alloc_scratch(1);
    lset(chunks, current, value, line);
    push_ip4_octets(chunks, current, value, line);
    chunks[current].emit_array_new_fixed(0, 4, line);
}

/// `_vybe_ip4_mask(prefixlen)` → the 32-bit netmask.
/// Stack: `[number] -> [number]`.
pub fn emit_ip4_mask(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_f64_const(0.0, line);
    }
    let bits = chunks[current].alloc_scratch(1);
    lset(chunks, current, bits, line);
    // 2^32 - 2^(32 - bits)
    chunks[current].emit_f64_const(4294967296.0, line);
    chunks[current].emit_f64_const(2.0, line);
    chunks[current].emit_f64_const(32.0, line);
    lget(chunks, current, bits, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    call_import(chunks, current, "ecma:math", "pow", 2, line);
    chunks[current].emit_op(Op::F64_SUB, line);
}

/// `_vybe_ip4_count(prefixlen)` → the address count, `2^(32 - prefixlen)`.
/// Stack: `[number] -> [number]`.
pub fn emit_ip4_count(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_f64_const(0.0, line);
    }
    let bits = chunks[current].alloc_scratch(1);
    lset(chunks, current, bits, line);
    chunks[current].emit_f64_const(2.0, line);
    chunks[current].emit_f64_const(32.0, line);
    lget(chunks, current, bits, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    call_import(chunks, current, "ecma:math", "pow", 2, line);
}

/// `_vybe_ip4_net_parts("192.168.1.0/24")` → `(address_int, prefixlen)`. A
/// missing `/len` defaults to /32, matching CPython.
/// Stack: `[str] -> [tuple]`.
pub fn emit_ip4_net_parts(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
    }
    let base = stash_args(chunks, current, argc, line);
    let parts = chunks[current].alloc_scratch(1);
    lget(chunks, current, base, line);
    chunks[current].emit_string_const("/", line);
    call_import(chunks, current, "ecma:string", "split", 2, line);
    lset(chunks, current, parts, line);

    lget(chunks, current, parts, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    emit_ip4_parse(chunks, current, 1, line);

    lget(chunks, current, parts, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if_value(line);
    lget(chunks, current, parts, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_f64_const(10.0, line);
    call_import(chunks, current, "ecma:number", "parseInt", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_f64_const(32.0, line);
    chunks[current].emit_end(line);

    tuples::emit_tuple(chunks, current, 2, line);
}
