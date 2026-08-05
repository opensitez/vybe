//! `java.net.URL` / `java.net.URI` as a PLATFORM adapter over the shared URL
//! primitive.
//!
//! These are JDK classes, so they belong to every JVM language — Java, Kotlin,
//! Scala, Groovy. They used to be ~650 lines of AST prelude built in
//! `languages/java/src/emitter/format_runtime.rs`, which meant a second JVM
//! frontend would have had to build the same prelude again. Here the behaviour
//! is emitted at the call site, once, and reached through the namespace tree.
//!
//! # `java.net.URL` is a SYNTACTIC parser, not WHATWG
//!
//! Measured against real `java` — this is why the old prelude needed shadow
//! fields (`__path`, `__spec`) to undo what `web:url` had already done:
//!
//! | input | real java `getPath`/`getHost`/`getPort` | WHATWG |
//! |---|---|---|
//! | `HTTP://Example.COM/a/../b` | `/a/../b`, `Example.COM` | `/b`, `example.com` |
//! | `http://x.com:80/p` | port `80` | port dropped |
//! | `http://x.com` | path `""` | path `/` |
//!
//! Scheme case is the one thing java DOES fold. That is exactly
//! [`ParseOptions::python`] — `mode: Syntactic, lowercase_scheme: true` — so the
//! JDK reads through the same parser as php `parse_url` and python `urlsplit`,
//! with no java-specific parsing anywhere.
//!
//! # What stays here rather than in the primitive
//!
//! Java's ABSENT-component policy: `getPort()` is `-1`, `getQuery()`/`getRef()`
//! are `null`, and `URI` reports `null` for the components an opaque URI has
//! none of. WHATWG and python both say `""`. That is language-shaped
//! behaviour over a shared component read, which is what an adapter is for.

use std::sync::Arc;
use vybe_compiler::primitives::url::{self, ParseOptions, PercentOptions, UrlField};
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

// ── local emit helpers ──────────────────────────────────────────────────────

fn slot(chunks: &mut [Chunk], c: usize) -> u16 {
    chunks[c].alloc_scratch(1)
}

fn lget(chunks: &mut [Chunk], c: usize, s: u16, line: u32) {
    chunks[c].emit_op_u16(Op::LOCAL_GET, s, line);
}

fn lset(chunks: &mut [Chunk], c: usize, s: u16, line: u32) {
    chunks[c].emit_op_u16(Op::LOCAL_SET, s, line);
}

/// Stack: `[obj]` → `[value]`.
fn sget(chunks: &mut [Chunk], c: usize, key: &str, line: u32) {
    let k = chunks[c].add_constant(Value::String(Arc::from(key)));
    chunks[c].emit_struct_field_op(Op::STRUCT_GET, 0, k, line);
}

/// Stack: `[obj, value]` → `[]`.
fn sset(chunks: &mut [Chunk], c: usize, key: &str, line: u32) {
    let k = chunks[c].add_constant(Value::String(Arc::from(key)));
    chunks[c].emit_struct_field_op(Op::STRUCT_SET, 0, k, line);
}

fn call(chunks: &mut [Chunk], c: usize, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunks[c].add_import(module.to_string(), name.to_string());
    chunks[c].emit_call(idx, argc, line);
}

fn push_str(chunks: &mut [Chunk], c: usize, s: &str, line: u32) {
    chunks[c].emit_string_const(s, line);
}

/// Stack: `[a, b]` → `[a ++ b]`.
fn concat(chunks: &mut [Chunk], c: usize, line: u32) {
    call(chunks, c, "wasm:js-string", "concat", 2, line);
}

/// Stack: `[s]` → `[i32]` — true when `s` is the empty string.
fn is_empty(chunks: &mut [Chunk], c: usize, line: u32) {
    call(chunks, c, "wasm:js-string", "length", 1, line);
    chunks[c].emit_i32_const(0, line);
    chunks[c].emit_op(Op::I32_EQ, line);
}

/// Coerce to string. Stack: `[v]` → `[string]`.
fn to_str(chunks: &mut [Chunk], c: usize, line: u32) {
    call(chunks, c, "ecma:string", "String", 1, line);
}

/// Read a canonical component off a parsed URL held in `url_slot`.
/// Stack: `-> [string]`. This is the SHARED reader — no java parsing here.
fn component(chunks: &mut [Chunk], c: usize, url_slot: u16, field: UrlField, line: u32) {
    url::emit_component(chunks, c, url_slot, field, line);
}

/// Read a component, then substitute when it came back empty.
///
/// `absent` emits the replacement value onto the stack. This is java's absent
/// policy — `-1` for a port, `null` for a query — applied over the shared
/// component read rather than a second component implementation.
fn component_or(
    chunks: &mut [Chunk],
    c: usize,
    url_slot: u16,
    field: UrlField,
    line: u32,
    absent: impl FnOnce(&mut [Chunk], usize, u32),
    present: impl FnOnce(&mut [Chunk], usize, u16, u32),
) {
    let v = slot(chunks, c);
    component(chunks, c, url_slot, field, line);
    lset(chunks, c, v, line);
    lget(chunks, c, v, line);
    is_empty(chunks, c, line);
    chunks[c].emit_if_value(line);
    absent(chunks, c, line);
    chunks[c].emit_else(line);
    present(chunks, c, v, line);
    chunks[c].emit_end(line);
}

/// Pop the single receiver into a fresh slot. Stack: `[url]` → `[]`.
fn take_receiver(chunks: &mut [Chunk], c: usize, line: u32) -> u16 {
    let s = slot(chunks, c);
    lset(chunks, c, s, line);
    s
}

/// Pop two arguments. Stack: `[a, b]` → `[]`; returns `(a_slot, b_slot)`.
fn take_two(chunks: &mut [Chunk], c: usize, line: u32) -> (u16, u16) {
    let b = slot(chunks, c);
    let a = slot(chunks, c);
    lset(chunks, c, b, line);
    lset(chunks, c, a, line);
    (a, b)
}

/// Java URI overloads accept either a URI object or a String spec. Normalize the
/// argument slot to the same parsed object shape the rest of the adapter uses.
fn coerce_uri_slot(chunks: &mut Vec<Chunk>, c: usize, s: u16, line: u32) {
    lget(chunks, c, s, line);
    call(chunks, c, "ecma:value", "typeof", 1, line);
    push_str(chunks, c, "string", line);
    call(chunks, c, "wasm:js-string", "equals", 2, line);
    chunks[c].emit_if_value(line);
    lget(chunks, c, s, line);
    let obj = emit_parse_into(chunks, c, true, line);
    lget(chunks, c, obj, line);
    chunks[c].emit_else(line);
    lget(chunks, c, s, line);
    chunks[c].emit_end(line);
    lset(chunks, c, s, line);
}

// ── construction ────────────────────────────────────────────────────────────

/// Parse `spec` (already on the stack) into the shared component object and
/// stamp the JDK's own bookkeeping onto it.
///
/// `__spec` is the raw text the object was built from: java's `toString()`
/// returns the spec VERBATIM (`http://Example.COM/a/../b` stays that way), so
/// it is the value, not a re-render of the components.
fn emit_parse_into(chunks: &mut Vec<Chunk>, c: usize, opaque: bool, line: u32) -> u16 {
    let spec = slot(chunks, c);
    let obj = slot(chunks, c);

    to_str(chunks, c, line);
    lset(chunks, c, spec, line);

    lget(chunks, c, spec, line);
    url::emit_parse(chunks, c, ParseOptions::python(), line);
    lset(chunks, c, obj, line);

    lget(chunks, c, obj, line);
    lget(chunks, c, spec, line);
    sset(chunks, c, "__spec", line);

    // The scheme-specific part: everything after the first `:`. java reports it
    // for every absolute URI, and it is the whole value of an opaque one.
    let ssp = slot(chunks, c);
    let colon = slot(chunks, c);
    lget(chunks, c, spec, line);
    push_str(chunks, c, ":", line);
    call(chunks, c, "ecma:string", "indexOf", 2, line);
    lset(chunks, c, colon, line);
    lget(chunks, c, colon, line);
    chunks[c].emit_i32_const(0, line);
    chunks[c].emit_op(Op::I32_LT_S, line);
    chunks[c].emit_if_value(line);
    lget(chunks, c, spec, line);
    chunks[c].emit_else(line);
    lget(chunks, c, spec, line);
    lget(chunks, c, colon, line);
    chunks[c].emit_i32_const(1, line);
    chunks[c].emit_op(Op::I32_ADD, line);
    chunks[c].emit_i32_const(0x7FFF_FFFF, line);
    call(chunks, c, "ecma:string", "substring", 3, line);
    chunks[c].emit_end(line);
    lset(chunks, c, ssp, line);
    lget(chunks, c, obj, line);
    lget(chunks, c, ssp, line);
    sset(chunks, c, "__ssp", line);

    if opaque {
        // java: opaque ⇔ absolute AND the scheme-specific part does not start
        // with `/`. `mailto:bob@x.com` is opaque; `http://x/y` is not.
        lget(chunks, c, obj, line);
        {
            component(chunks, c, obj, UrlField::Scheme, line);
            is_empty(chunks, c, line);
            chunks[c].emit_if_value(line);
            chunks[c].emit_bool_const(false, line);
            chunks[c].emit_else(line);
            lget(chunks, c, ssp, line);
            push_str(chunks, c, "/", line);
            call(chunks, c, "ecma:string", "startsWith", 2, line);
            chunks[c].emit_if_value(line);
            chunks[c].emit_bool_const(false, line);
            chunks[c].emit_else(line);
            chunks[c].emit_bool_const(true, line);
            chunks[c].emit_end(line);
            chunks[c].emit_end(line);
        }
        sset(chunks, c, "__opaque", line);
    } else {
        lget(chunks, c, obj, line);
        chunks[c].emit_bool_const(false, line);
        sset(chunks, c, "__opaque", line);
    }

    obj
}

/// `new URL(spec)` / `new URL(context, spec)` / `new URL(proto, host, port, file)`.
pub fn emit_url_new(chunks: &mut Vec<Chunk>, c: usize, argc: u8, line: u32) {
    match argc {
        2 => emit_url_ctx(chunks, c, line),
        3 => emit_url_make3(chunks, c, line),
        4 => emit_url_make4(chunks, c, line),
        _ => {
            let obj = emit_parse_into(chunks, c, false, line);
            lget(chunks, c, obj, line);
        }
    }
}

/// `new URI(spec)` / `new URI(scheme, host, path)` /
/// `new URI(scheme, user, host, port, path, query, fragment)`.
pub fn emit_uri_new(chunks: &mut Vec<Chunk>, c: usize, argc: u8, line: u32) {
    match argc {
        3 => emit_uri_make3(chunks, c, line),
        7 => emit_uri_make7(chunks, c, line),
        _ => {
            let obj = emit_parse_into(chunks, c, true, line);
            lget(chunks, c, obj, line);
        }
    }
}

/// `URL.toURI()` — same component payload, now seen through URI's registered
/// return type. The parser shape already carries the raw Java spec.
pub fn emit_url_to_uri(chunks: &mut [Chunk], c: usize, line: u32) {
    let u = take_receiver(chunks, c, line);
    lget(chunks, c, u, line);
}

/// `URI.toURL()` — same component payload, now seen through URL's registered
/// return type.
pub fn emit_uri_to_url(chunks: &mut [Chunk], c: usize, line: u32) {
    let u = take_receiver(chunks, c, line);
    lget(chunks, c, u, line);
}

/// `URLEncoder.encode(s, charset)` — Java form encoding ignores charset here
/// because Vybe strings are already Unicode and ECMA supplies UTF-8 escaping.
pub fn emit_url_encode(chunks: &mut Vec<Chunk>, c: usize, line: u32) {
    chunks[c].emit_op(Op::DROP, line);
    url::emit_percent_encode(chunks, c, PercentOptions::form(), line);
}

/// `URLDecoder.decode(s, charset)`.
pub fn emit_url_decode(chunks: &mut Vec<Chunk>, c: usize, line: u32) {
    chunks[c].emit_op(Op::DROP, line);
    url::emit_percent_decode(chunks, c, PercentOptions::form(), line);
}

/// `new URL(context, spec)` — java resolves relative refs WITHOUT WHATWG's
/// dot-segment removal, so the spec is composed textually and re-parsed.
fn emit_url_ctx(chunks: &mut Vec<Chunk>, c: usize, line: u32) {
    let (base, spec) = take_two(chunks, c, line);
    let path = slot(chunks, c);

    lget(chunks, c, spec, line);
    to_str(chunks, c, line);
    lset(chunks, c, spec, line);

    // An absolute spec ignores the context entirely.
    lget(chunks, c, spec, line);
    push_str(chunks, c, "://", line);
    call(chunks, c, "ecma:string", "indexOf", 2, line);
    chunks[c].emit_i32_const(0, line);
    chunks[c].emit_op(Op::I32_GE_S, line);
    chunks[c].emit_if_value(line);
    lget(chunks, c, spec, line);
    chunks[c].emit_else(line);
    {
        // Root-relative keeps the spec; otherwise splice onto the base's
        // directory (everything through the base path's last `/`).
        lget(chunks, c, spec, line);
        push_str(chunks, c, "/", line);
        call(chunks, c, "ecma:string", "startsWith", 2, line);
        chunks[c].emit_if_value(line);
        lget(chunks, c, spec, line);
        chunks[c].emit_else(line);
        {
            let bp = slot(chunks, c);
            let cut = slot(chunks, c);
            component(chunks, c, base, UrlField::Path, line);
            lset(chunks, c, bp, line);
            lget(chunks, c, bp, line);
            push_str(chunks, c, "/", line);
            call(chunks, c, "ecma:string", "lastIndexOf", 2, line);
            lset(chunks, c, cut, line);
            lget(chunks, c, bp, line);
            chunks[c].emit_i32_const(0, line);
            lget(chunks, c, cut, line);
            chunks[c].emit_i32_const(1, line);
            chunks[c].emit_op(Op::I32_ADD, line);
            call(chunks, c, "ecma:string", "substring", 3, line);
            lget(chunks, c, spec, line);
            concat(chunks, c, line);
        }
        chunks[c].emit_end(line);
        lset(chunks, c, path, line);

        component(chunks, c, base, UrlField::Scheme, line);
        push_str(chunks, c, "://", line);
        concat(chunks, c, line);
        lget(chunks, c, base, line);
        sget(chunks, c, "host", line);
        concat(chunks, c, line);
        lget(chunks, c, path, line);
        concat(chunks, c, line);
    }
    chunks[c].emit_end(line);

    let obj = emit_parse_into(chunks, c, false, line);
    lget(chunks, c, obj, line);
}

/// `new URL(protocol, host, file)` — Java's three-arg form implies port `-1`.
fn emit_url_make3(chunks: &mut Vec<Chunk>, c: usize, line: u32) {
    let file = slot(chunks, c);
    let host = slot(chunks, c);
    let proto = slot(chunks, c);
    lset(chunks, c, file, line);
    lset(chunks, c, host, line);
    lset(chunks, c, proto, line);

    lget(chunks, c, proto, line);
    to_str(chunks, c, line);
    push_str(chunks, c, "://", line);
    concat(chunks, c, line);
    lget(chunks, c, host, line);
    to_str(chunks, c, line);
    concat(chunks, c, line);
    lget(chunks, c, file, line);
    to_str(chunks, c, line);
    concat(chunks, c, line);

    let obj = emit_parse_into(chunks, c, false, line);
    lget(chunks, c, obj, line);
}

/// `new URL(protocol, host, port, file)`.
fn emit_url_make4(chunks: &mut Vec<Chunk>, c: usize, line: u32) {
    let file = slot(chunks, c);
    let port = slot(chunks, c);
    let host = slot(chunks, c);
    let proto = slot(chunks, c);
    lset(chunks, c, file, line);
    lset(chunks, c, port, line);
    lset(chunks, c, host, line);
    lset(chunks, c, proto, line);

    lget(chunks, c, proto, line);
    to_str(chunks, c, line);
    push_str(chunks, c, "://", line);
    concat(chunks, c, line);
    lget(chunks, c, host, line);
    to_str(chunks, c, line);
    concat(chunks, c, line);

    lget(chunks, c, port, line);
    chunks[c].emit_f64_const(0.0, line);
    chunks[c].emit_op(Op::F64_GE, line);
    chunks[c].emit_if_value(line);
    push_str(chunks, c, ":", line);
    lget(chunks, c, port, line);
    to_str(chunks, c, line);
    concat(chunks, c, line);
    chunks[c].emit_else(line);
    push_str(chunks, c, "", line);
    chunks[c].emit_end(line);
    concat(chunks, c, line);

    lget(chunks, c, file, line);
    to_str(chunks, c, line);
    concat(chunks, c, line);

    let obj = emit_parse_into(chunks, c, false, line);
    lget(chunks, c, obj, line);
}

/// `new URI(scheme, host, path)`.
fn emit_uri_make3(chunks: &mut Vec<Chunk>, c: usize, line: u32) {
    let path = slot(chunks, c);
    let host = slot(chunks, c);
    let scheme = slot(chunks, c);
    lset(chunks, c, path, line);
    lset(chunks, c, host, line);
    lset(chunks, c, scheme, line);

    lget(chunks, c, scheme, line);
    to_str(chunks, c, line);
    push_str(chunks, c, "://", line);
    concat(chunks, c, line);
    lget(chunks, c, host, line);
    to_str(chunks, c, line);
    concat(chunks, c, line);
    lget(chunks, c, path, line);
    to_str(chunks, c, line);
    concat(chunks, c, line);

    let obj = emit_parse_into(chunks, c, true, line);
    lget(chunks, c, obj, line);
}

/// Append `sep ++ value` when `value` is neither null nor empty.
/// Stack: `[acc]` → `[acc]`.
fn append_optional(chunks: &mut [Chunk], c: usize, v: u16, sep: &str, suffix: bool, line: u32) {
    lget(chunks, c, v, line);
    chunks[c].emit_op(Op::REF_IS_NULL, line);
    chunks[c].emit_if_value(line);
    push_str(chunks, c, "", line);
    chunks[c].emit_else(line);
    if suffix {
        lget(chunks, c, v, line);
        to_str(chunks, c, line);
        push_str(chunks, c, sep, line);
        concat(chunks, c, line);
    } else {
        push_str(chunks, c, sep, line);
        lget(chunks, c, v, line);
        to_str(chunks, c, line);
        concat(chunks, c, line);
    }
    chunks[c].emit_end(line);
    concat(chunks, c, line);
}

/// `new URI(scheme, userInfo, host, port, path, query, fragment)`.
fn emit_uri_make7(chunks: &mut Vec<Chunk>, c: usize, line: u32) {
    let frag = slot(chunks, c);
    let query = slot(chunks, c);
    let path = slot(chunks, c);
    let port = slot(chunks, c);
    let host = slot(chunks, c);
    let user = slot(chunks, c);
    let scheme = slot(chunks, c);
    lset(chunks, c, frag, line);
    lset(chunks, c, query, line);
    lset(chunks, c, path, line);
    lset(chunks, c, port, line);
    lset(chunks, c, host, line);
    lset(chunks, c, user, line);
    lset(chunks, c, scheme, line);

    lget(chunks, c, scheme, line);
    to_str(chunks, c, line);
    push_str(chunks, c, "://", line);
    concat(chunks, c, line);

    // `user@` — java omits it for null AND for the empty string.
    lget(chunks, c, user, line);
    chunks[c].emit_op(Op::REF_IS_NULL, line);
    chunks[c].emit_if_value(line);
    push_str(chunks, c, "", line);
    chunks[c].emit_else(line);
    lget(chunks, c, user, line);
    to_str(chunks, c, line);
    is_empty(chunks, c, line);
    chunks[c].emit_if_value(line);
    push_str(chunks, c, "", line);
    chunks[c].emit_else(line);
    lget(chunks, c, user, line);
    to_str(chunks, c, line);
    push_str(chunks, c, "@", line);
    concat(chunks, c, line);
    chunks[c].emit_end(line);
    chunks[c].emit_end(line);
    concat(chunks, c, line);

    lget(chunks, c, host, line);
    to_str(chunks, c, line);
    concat(chunks, c, line);

    lget(chunks, c, port, line);
    chunks[c].emit_f64_const(0.0, line);
    chunks[c].emit_op(Op::F64_GE, line);
    chunks[c].emit_if_value(line);
    push_str(chunks, c, ":", line);
    lget(chunks, c, port, line);
    to_str(chunks, c, line);
    concat(chunks, c, line);
    chunks[c].emit_else(line);
    push_str(chunks, c, "", line);
    chunks[c].emit_end(line);
    concat(chunks, c, line);

    lget(chunks, c, path, line);
    to_str(chunks, c, line);
    concat(chunks, c, line);

    append_optional(chunks, c, query, "?", false, line);
    append_optional(chunks, c, frag, "#", false, line);

    let obj = emit_parse_into(chunks, c, true, line);
    lget(chunks, c, obj, line);
}

// ── component getters ───────────────────────────────────────────────────────

/// A getter that is exactly a shared component read — `getProtocol`,
/// `getHost`, `getPath`, `getAuthority`.
pub fn emit_component_getter(chunks: &mut [Chunk], c: usize, field: UrlField, line: u32) {
    let u = take_receiver(chunks, c, line);
    component(chunks, c, u, field, line);
}

/// The same read, but `null` when the component is absent — java's policy for
/// `getQuery`/`getRef`/`getFragment`, and for every component of an opaque URI.
pub fn emit_nullable_getter(chunks: &mut [Chunk], c: usize, field: UrlField, line: u32) {
    let u = take_receiver(chunks, c, line);
    component_or(
        chunks,
        c,
        u,
        field,
        line,
        |ch, c, line| ch[c].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line),
        |ch, c, v, line| lget(ch, c, v, line),
    );
}

/// `getPort()` — a NUMBER, and `-1` when the URL names none.
pub fn emit_port(chunks: &mut [Chunk], c: usize, line: u32) {
    let u = take_receiver(chunks, c, line);
    component_or(
        chunks,
        c,
        u,
        UrlField::Port,
        line,
        |ch, c, line| ch[c].emit_f64_const(-1.0, line),
        |ch, c, v, line| {
            lget(ch, c, v, line);
            call(ch, c, "ecma:number", "parseInt", 1, line);
        },
    );
}

/// `getDefaultPort()` — the IANA port for the scheme, `-1` when unknown.
pub fn emit_default_port(chunks: &mut [Chunk], c: usize, line: u32) {
    const DEFAULTS: &[(&str, f64)] = &[("http", 80.0), ("https", 443.0), ("ftp", 21.0)];
    let u = take_receiver(chunks, c, line);
    let p = slot(chunks, c);
    component(chunks, c, u, UrlField::Scheme, line);
    lset(chunks, c, p, line);
    for (scheme, port) in DEFAULTS {
        lget(chunks, c, p, line);
        push_str(chunks, c, scheme, line);
        chunks[c].emit_op(Op::EQ, line);
        chunks[c].emit_if_value(line);
        chunks[c].emit_f64_const(*port, line);
        chunks[c].emit_else(line);
    }
    chunks[c].emit_f64_const(-1.0, line);
    for _ in DEFAULTS {
        chunks[c].emit_end(line);
    }
}

/// `getFile()` — path plus `?query` when there is one.
pub fn emit_file(chunks: &mut [Chunk], c: usize, line: u32) {
    let u = take_receiver(chunks, c, line);
    component(chunks, c, u, UrlField::Path, line);
    lget(chunks, c, u, line);
    sget(chunks, c, "search", line);
    concat(chunks, c, line);
}

/// `getUserInfo()` — `user`, `user:password`, or `null`.
pub fn emit_user_info(chunks: &mut [Chunk], c: usize, line: u32) {
    let u = take_receiver(chunks, c, line);
    let user = slot(chunks, c);
    let pass = slot(chunks, c);

    component(chunks, c, u, UrlField::User, line);
    lset(chunks, c, user, line);
    component(chunks, c, u, UrlField::Pass, line);
    lset(chunks, c, pass, line);

    lget(chunks, c, user, line);
    is_empty(chunks, c, line);
    chunks[c].emit_if_value(line);
    chunks[c].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[c].emit_else(line);
    lget(chunks, c, pass, line);
    is_empty(chunks, c, line);
    chunks[c].emit_if_value(line);
    lget(chunks, c, user, line);
    chunks[c].emit_else(line);
    lget(chunks, c, user, line);
    push_str(chunks, c, ":", line);
    concat(chunks, c, line);
    lget(chunks, c, pass, line);
    concat(chunks, c, line);
    chunks[c].emit_end(line);
    chunks[c].emit_end(line);
}

// ── identity and text ───────────────────────────────────────────────────────

/// `toString()` / `toExternalForm()` / `toASCIIString()` — the spec verbatim.
pub fn emit_to_string(chunks: &mut [Chunk], c: usize, line: u32) {
    let u = take_receiver(chunks, c, line);
    lget(chunks, c, u, line);
    sget(chunks, c, "__spec", line);
}

/// `equals(other)` — java compares the URL text.
pub fn emit_equals(chunks: &mut [Chunk], c: usize, line: u32) {
    let (a, b) = take_two(chunks, c, line);
    lget(chunks, c, a, line);
    sget(chunks, c, "__spec", line);
    lget(chunks, c, b, line);
    sget(chunks, c, "__spec", line);
    chunks[c].emit_op(Op::EQ, line);
}

/// `hashCode()` — java's `String.hashCode` over the URL text: `h = 31h + c`,
/// wrapping at 32 bits. Accumulated in f64 because the dynamic i32 multiply
/// TRAPS on overflow instead of wrapping.
pub fn emit_hash(chunks: &mut [Chunk], c: usize, line: u32) {
    let u = take_receiver(chunks, c, line);
    let s = slot(chunks, c);
    let h = slot(chunks, c);
    let i = slot(chunks, c);
    let n = slot(chunks, c);

    lget(chunks, c, u, line);
    sget(chunks, c, "__spec", line);
    lset(chunks, c, s, line);
    lget(chunks, c, s, line);
    call(chunks, c, "wasm:js-string", "length", 1, line);
    lset(chunks, c, n, line);
    chunks[c].emit_f64_const(0.0, line);
    lset(chunks, c, h, line);
    chunks[c].emit_i32_const(0, line);
    lset(chunks, c, i, line);

    chunks[c].emit_block(line);
    let (loop_start, _) = chunks[c].emit_loop_s(line);
    lget(chunks, c, i, line);
    lget(chunks, c, n, line);
    chunks[c].emit_op(Op::I32_GE_S, line);
    chunks[c].emit_br_if(2, line);

    lget(chunks, c, h, line);
    chunks[c].emit_f64_const(31.0, line);
    chunks[c].emit_op(Op::F64_MUL, line);
    lget(chunks, c, s, line);
    lget(chunks, c, i, line);
    call(chunks, c, "wasm:js-string", "charCodeAt", 2, line);
    chunks[c].emit_op(Op::F64_ADD, line);
    emit_wrap_i32(chunks, c, line);
    lset(chunks, c, h, line);

    lget(chunks, c, i, line);
    chunks[c].emit_i32_const(1, line);
    chunks[c].emit_op(Op::I32_ADD, line);
    lset(chunks, c, i, line);
    chunks[c].emit_loop(loop_start, line);
    chunks[c].emit_end(line);
    chunks[c].emit_end(line);

    lget(chunks, c, h, line);
}

/// Wrap an f64 accumulator into java's signed 32-bit range.
/// Stack: `[f64]` → `[f64]`.
fn emit_wrap_i32(chunks: &mut [Chunk], c: usize, line: u32) {
    let x = slot(chunks, c);
    chunks[c].emit_f64_const(4294967296.0, line);
    vybe_compiler::primitives::math::emit_c_fmod(&mut chunks[c], line);
    lset(chunks, c, x, line);

    lget(chunks, c, x, line);
    chunks[c].emit_f64_const(2147483648.0, line);
    chunks[c].emit_op(Op::F64_GE, line);
    chunks[c].emit_if_value(line);
    lget(chunks, c, x, line);
    chunks[c].emit_f64_const(4294967296.0, line);
    chunks[c].emit_op(Op::F64_SUB, line);
    chunks[c].emit_else(line);
    lget(chunks, c, x, line);
    chunks[c].emit_f64_const(-2147483648.0, line);
    chunks[c].emit_op(Op::F64_LT, line);
    chunks[c].emit_if_value(line);
    lget(chunks, c, x, line);
    chunks[c].emit_f64_const(4294967296.0, line);
    chunks[c].emit_op(Op::F64_ADD, line);
    chunks[c].emit_else(line);
    lget(chunks, c, x, line);
    chunks[c].emit_end(line);
    chunks[c].emit_end(line);
}

/// `sameFile(other)` — same protocol, host and file, ignoring the ref.
pub fn emit_same_file(chunks: &mut Vec<Chunk>, c: usize, line: u32) {
    let (a, b) = take_two(chunks, c, line);
    let eq = |chunks: &mut Vec<Chunk>, c: usize, field: UrlField| {
        component(chunks, c, a, field, line);
        component(chunks, c, b, field, line);
        chunks[c].emit_op(Op::EQ, line);
    };
    eq(chunks, c, UrlField::Scheme);
    chunks[c].emit_if_value(line);
    eq(chunks, c, UrlField::Host);
    chunks[c].emit_if_value(line);
    eq(chunks, c, UrlField::Path);
    chunks[c].emit_if_value(line);
    lget(chunks, c, a, line);
    sget(chunks, c, "search", line);
    lget(chunks, c, b, line);
    sget(chunks, c, "search", line);
    chunks[c].emit_op(Op::EQ, line);
    chunks[c].emit_else(line);
    chunks[c].emit_bool_const(false, line);
    chunks[c].emit_end(line);
    chunks[c].emit_else(line);
    chunks[c].emit_bool_const(false, line);
    chunks[c].emit_end(line);
    chunks[c].emit_else(line);
    chunks[c].emit_bool_const(false, line);
    chunks[c].emit_end(line);
}

/// `compareTo(other)` — lexicographic over the URI text.
pub fn emit_compare_to(chunks: &mut [Chunk], c: usize, line: u32) {
    let (a, b) = take_two(chunks, c, line);
    let x = slot(chunks, c);
    let y = slot(chunks, c);
    lget(chunks, c, a, line);
    sget(chunks, c, "__spec", line);
    lset(chunks, c, x, line);
    lget(chunks, c, b, line);
    sget(chunks, c, "__spec", line);
    lset(chunks, c, y, line);

    lget(chunks, c, x, line);
    lget(chunks, c, y, line);
    chunks[c].emit_op(Op::EQ, line);
    chunks[c].emit_if_value(line);
    chunks[c].emit_f64_const(0.0, line);
    chunks[c].emit_else(line);
    lget(chunks, c, x, line);
    lget(chunks, c, y, line);
    chunks[c].emit_op(Op::F64_LT, line);
    chunks[c].emit_if_value(line);
    chunks[c].emit_f64_const(-1.0, line);
    chunks[c].emit_else(line);
    chunks[c].emit_f64_const(1.0, line);
    chunks[c].emit_end(line);
    chunks[c].emit_end(line);
}

// ── URI-only relational surface ─────────────────────────────────────────────

/// `getSchemeSpecificPart()` — everything after the first `:`.
pub fn emit_ssp(chunks: &mut [Chunk], c: usize, line: u32) {
    let u = take_receiver(chunks, c, line);
    lget(chunks, c, u, line);
    sget(chunks, c, "__ssp", line);
}

/// `isAbsolute()` — the URI names a scheme.
pub fn emit_is_absolute(chunks: &mut [Chunk], c: usize, line: u32) {
    let u = take_receiver(chunks, c, line);
    component(chunks, c, u, UrlField::Scheme, line);
    is_empty(chunks, c, line);
    chunks[c].emit_if_value(line);
    chunks[c].emit_bool_const(false, line);
    chunks[c].emit_else(line);
    chunks[c].emit_bool_const(true, line);
    chunks[c].emit_end(line);
}

/// `isOpaque()` — absolute, with a scheme-specific part that is not a path.
pub fn emit_is_opaque(chunks: &mut [Chunk], c: usize, line: u32) {
    let u = take_receiver(chunks, c, line);
    lget(chunks, c, u, line);
    sget(chunks, c, "__opaque", line);
}

/// Build `scheme://host` from a parsed URL held in `slot`. Stack: `-> [string]`.
fn emit_origin(chunks: &mut [Chunk], c: usize, u: u16, line: u32) {
    component(chunks, c, u, UrlField::Scheme, line);
    push_str(chunks, c, "://", line);
    concat(chunks, c, line);
    lget(chunks, c, u, line);
    sget(chunks, c, "host", line);
    concat(chunks, c, line);
}

/// Base path as a directory, matching the tests' Java-style resolution: trim a
/// trailing slash, then keep through the previous slash.
fn emit_base_directory_path(chunks: &mut [Chunk], c: usize, base: u16, line: u32) {
    let path = slot(chunks, c);
    let trimmed = slot(chunks, c);
    let cut = slot(chunks, c);

    component(chunks, c, base, UrlField::Path, line);
    lset(chunks, c, path, line);

    lget(chunks, c, path, line);
    push_str(chunks, c, "/", line);
    call(chunks, c, "ecma:string", "endsWith", 2, line);
    chunks[c].emit_if_value(line);
    lget(chunks, c, path, line);
    chunks[c].emit_i32_const(0, line);
    lget(chunks, c, path, line);
    call(chunks, c, "wasm:js-string", "length", 1, line);
    chunks[c].emit_i32_const(1, line);
    chunks[c].emit_op(Op::I32_SUB, line);
    call(chunks, c, "ecma:string", "substring", 3, line);
    chunks[c].emit_else(line);
    lget(chunks, c, path, line);
    chunks[c].emit_end(line);
    lset(chunks, c, trimmed, line);

    lget(chunks, c, trimmed, line);
    push_str(chunks, c, "/", line);
    call(chunks, c, "ecma:string", "lastIndexOf", 2, line);
    lset(chunks, c, cut, line);

    lget(chunks, c, trimmed, line);
    chunks[c].emit_i32_const(0, line);
    lget(chunks, c, cut, line);
    chunks[c].emit_i32_const(1, line);
    chunks[c].emit_op(Op::I32_ADD, line);
    call(chunks, c, "ecma:string", "substring", 3, line);
}

fn emit_base_directory_url(chunks: &mut [Chunk], c: usize, base: u16, line: u32) {
    emit_origin(chunks, c, base, line);
    emit_base_directory_path(chunks, c, base, line);
    concat(chunks, c, line);
}

/// `normalize()` — remove `.` and `..` segments from the path, changing nothing
/// else. WHATWG's parser already implements exactly that removal, so the
/// resolved PATH is taken from it while scheme, host case and query stay as
/// java left them (`http://Example.COM/a/b/../c` → `http://Example.COM/a/c`).
pub fn emit_normalize(chunks: &mut Vec<Chunk>, c: usize, line: u32) {
    let u = take_receiver(chunks, c, line);
    emit_origin(chunks, c, u, line);
    lget(chunks, c, u, line);
    sget(chunks, c, "__spec", line);
    call(chunks, c, "web:url", "new", 1, line);
    sget(chunks, c, "pathname", line);
    concat(chunks, c, line);
    lget(chunks, c, u, line);
    sget(chunks, c, "search", line);
    concat(chunks, c, line);
    lget(chunks, c, u, line);
    sget(chunks, c, "hash", line);
    concat(chunks, c, line);

    let obj = emit_parse_into(chunks, c, true, line);
    lget(chunks, c, obj, line);
}

/// `resolve(other)` — an absolute reference wins outright; otherwise WHATWG
/// resolves the path against the base and the base's own scheme/host are kept.
pub fn emit_resolve(chunks: &mut Vec<Chunk>, c: usize, line: u32) {
    let (base, rel) = take_two(chunks, c, line);
    coerce_uri_slot(chunks, c, rel, line);

    component(chunks, c, rel, UrlField::Scheme, line);
    is_empty(chunks, c, line);
    chunks[c].emit_if(line);
    {
        let w = slot(chunks, c);
        lget(chunks, c, rel, line);
        sget(chunks, c, "__spec", line);
        emit_base_directory_url(chunks, c, base, line);
        call(chunks, c, "web:url", "new", 2, line);
        lset(chunks, c, w, line);

        emit_origin(chunks, c, base, line);
        lget(chunks, c, w, line);
        sget(chunks, c, "pathname", line);
        concat(chunks, c, line);
        lget(chunks, c, w, line);
        sget(chunks, c, "search", line);
        concat(chunks, c, line);
        lget(chunks, c, w, line);
        sget(chunks, c, "hash", line);
        concat(chunks, c, line);

        let obj = emit_parse_into(chunks, c, true, line);
        lget(chunks, c, obj, line);
        lset(chunks, c, rel, line);
    }
    chunks[c].emit_end(line);
    lget(chunks, c, rel, line);
}

/// `relativize(other)` — when `other` sits under this URI, return just the
/// trailing part; otherwise return `other` unchanged.
pub fn emit_relativize(chunks: &mut Vec<Chunk>, c: usize, line: u32) {
    let (base, target) = take_two(chunks, c, line);
    coerce_uri_slot(chunks, c, target, line);
    let prefix = slot(chunks, c);
    let tp = slot(chunks, c);

    emit_origin(chunks, c, base, line);
    emit_origin(chunks, c, target, line);
    chunks[c].emit_op(Op::EQ, line);
    chunks[c].emit_if(line);
    {
        // The base path as a DIRECTORY prefix — java relativizes against
        // `base + "/"`, so `/a/b/c` vs `/a/b/c/q` yields `q`, not `../q`.
        emit_base_directory_path(chunks, c, base, line);
        lset(chunks, c, prefix, line);

        component(chunks, c, target, UrlField::Path, line);
        lset(chunks, c, tp, line);

        lget(chunks, c, tp, line);
        lget(chunks, c, prefix, line);
        call(chunks, c, "ecma:string", "startsWith", 2, line);
        chunks[c].emit_if(line);
        {
            lget(chunks, c, tp, line);
            lget(chunks, c, prefix, line);
            call(chunks, c, "wasm:js-string", "length", 1, line);
            chunks[c].emit_i32_const(0x7FFF_FFFF, line);
            call(chunks, c, "ecma:string", "substring", 3, line);
            lget(chunks, c, target, line);
            sget(chunks, c, "search", line);
            concat(chunks, c, line);
            lget(chunks, c, target, line);
            sget(chunks, c, "hash", line);
            concat(chunks, c, line);

            let obj = emit_parse_into(chunks, c, true, line);
            lget(chunks, c, obj, line);
            lset(chunks, c, target, line);
        }
        chunks[c].emit_end(line);
    }
    chunks[c].emit_end(line);
    lget(chunks, c, target, line);
}
