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

fn call_import(chunks: &mut [Chunk], current: usize, module: &str, name: &str, argc: u8, line: u32) {
    // Register on the CURRENT chunk — an index taken from chunks[0] resolves
    // to the wrong host fn when the code runs inside a function chunk.
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(argc, line);
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
    chunks[current].emit_op_u16(Op::ARRAY_NEW_FIXED, 4, line);
    // Same constructor `int.to_bytes` / the `b"..."` literal path uses, so
    // `len()` and indexing behave identically to any other `bytes` value.
    call_import(chunks, current, "ecma:uint8array", "new", 1, line);
}

/// `inet_ntoa(b"\xc0\xa8\x01\x01")` → `"192.168.1.1"`.
/// Stack: `[bytes] -> [str]`.
pub fn emit_inet_ntoa(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_op(Op::NULL, line);
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
        call_import(chunks, current, "wasi:sockets/instance-network", "instance-network", 0, line);
        lget(chunks, current, host, line);
        call_import(
            chunks,
            current,
            "wasi:sockets/ip-name-lookup",
            "resolve-addresses",
            2,
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
    chunks[current].emit_op_u16(Op::ARRAY_NEW_FIXED, 1, line);
}
