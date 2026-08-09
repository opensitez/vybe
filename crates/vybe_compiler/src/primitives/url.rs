//! `url` — the URL domain as adapter primitives.
//!
//! A URL is a STRUCTURED format, not a bare encoding, so the whole domain lives
//! together: percent-encoding, splitting/joining, query strings. ECMA-262
//! covers only `encodeURIComponent`/`decodeURIComponent`; everything above that
//! every language wrote itself — PHP's `parse_url` is a WALKER REWRITE, Python
//! has a 532-line `url_adapter.rs`. Same components, three surface shapes
//! (PHP assoc array, Python named tuple, JS `URL` object).
//!
//! # Percent-encoding is ONE algorithm with three parameters
//!
//! Measured against real `php` and `python3` on `"a b&c=d/é~"`:
//!
//! | | space | `~` | `/` |
//! |---|---|---|---|
//! | php `urlencode` | `+` | `%7E` | `%2F` |
//! | py `quote_plus` | `+` | `~` | `%2F` |
//! | php `rawurlencode` | `%20` | `~` | `%2F` |
//! | py `quote` | `%20` | `~` | `/` |
//!
//! Four bindings, one implementation — see [`PercentOptions`]. PHP escapes `~`
//! because `urlencode` predates RFC 3986; Python's `quote` leaves `/` safe
//! because it is meant for path segments.
//!
//! **No coercion here.** Every function takes STRINGS on the stack.

use std::sync::Arc;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

/// How a percent-encoder treats the characters the four variants disagree on.
#[derive(Clone, Copy)]
pub struct PercentOptions {
    /// Space as `+` (php `urlencode`, py `quote_plus`) rather than `%20`.
    /// On decode, whether `+` reads back as a space.
    pub space_as_plus: bool,
    /// Escape `~`. RFC 3986 lists it unreserved; php `urlencode` predates that.
    pub escape_tilde: bool,
    /// Left unescaped beyond the unreserved set — py `quote` defaults to `"/"`.
    pub safe: &'static str,
}

impl PercentOptions {
    /// php `urlencode`
    pub const fn form() -> PercentOptions {
        PercentOptions {
            space_as_plus: true,
            escape_tilde: true,
            safe: "",
        }
    }
    /// php `rawurlencode` — RFC 3986
    pub const fn rfc3986() -> PercentOptions {
        PercentOptions {
            space_as_plus: false,
            escape_tilde: false,
            safe: "",
        }
    }
    /// py `quote_plus`
    pub const fn form_rfc3986() -> PercentOptions {
        PercentOptions {
            space_as_plus: true,
            escape_tilde: false,
            safe: "",
        }
    }
    /// py `quote` — path-safe
    pub const fn path() -> PercentOptions {
        PercentOptions {
            space_as_plus: false,
            escape_tilde: false,
            safe: "/",
        }
    }
}

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::F64(v) => chunk.emit_f64_const(*v, line),
        Value::I32(v) => chunk.emit_i32_const(*v, line),
        Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line),
        Value::BigInt(v) => chunk.emit_i64_const(v.to_i64_wrapping(), line),
        Value::String(s) => chunk.emit_string_const(&s, line),
        Value::Bool(b) => chunk.emit_bool_const(*b, line),

        _ => {
            unreachable!("push_const: unexpected value type");
        }
    }
}

fn push_str(chunk: &mut Chunk, v: &str, line: u32) {
    push_const(chunk, Value::String(Arc::from(v)), line);
}

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module.to_string(), name.to_string());
    let chunk = &mut chunks[current];
    chunk.emit_call(idx, argc, line);
}

/// Percent-encode. Stack: `[s]` → `[string]`.
///
/// Built on `ecma:string.encodeURIComponent`, which is ECMA-262 §19.2.6.5 and
/// already does the hard part: UTF-8 bytes, uppercase hex, and the RFC 3986
/// unreserved set (`A-Za-z0-9 -_.!~*\'()`). Every variant below is that plus at
/// most two literal fixups, which is why they share one emitter.
pub fn emit_percent_encode(chunks: &mut [Chunk], current: usize, opts: PercentOptions, line: u32) {
    call_import(
        chunks,
        current,
        "ecma:string",
        "encodeURIComponent",
        1,
        line,
    );
    // `encodeURIComponent` leaves `~` unescaped; php `urlencode` predates
    // RFC 3986 and escapes it, so that variant asks for it back.
    if opts.escape_tilde {
        replace_all(chunks, current, "~", "%7E", line);
    }
    if opts.space_as_plus {
        replace_all(chunks, current, "%20", "+", line);
    }
    // Anything the caller wants left literal — py `quote` keeps `/`.
    for ch in opts.safe.chars() {
        let enc = format!("%{:02X}", ch as u32);
        replace_all(chunks, current, &enc, &ch.to_string(), line);
    }
}

/// Percent-decode. Stack: `[s]` → `[string]`.
///
/// `+` means a space only in the form-encoded variants (php `urldecode`, py
/// `unquote_plus`); the RFC 3986 variants leave it literal.
pub fn emit_percent_decode(chunks: &mut [Chunk], current: usize, opts: PercentOptions, line: u32) {
    if opts.space_as_plus {
        replace_all(chunks, current, "+", "%20", line);
    }
    call_import(
        chunks,
        current,
        "ecma:string",
        "decodeURIComponent",
        1,
        line,
    );
}

/// `s.replaceAll(from, to)` with literal (non-regex) operands.
fn replace_all(chunks: &mut [Chunk], current: usize, from: &str, to: &str, line: u32) {
    let chunk = &mut chunks[current];
    push_str(chunk, from, line);
    push_str(chunk, to, line);
    let idx = chunk.add_import("ecma:string", "replaceAll");
    chunk.emit_call(idx, 3, line);
}

// ── components: the normalized superset ───────────────────────────────
//
// `web:url` (WHATWG) is the parser and the storage. This is the SUPERSET over
// it: the fields WHATWG spells differently or does not expose at all, so one
// parsed URL reads the same from php, python and js.
//
// | canonical | php `parse_url` | python `SplitResult` | WHATWG `URL` |
// |---|---|---|---|
// | `Scheme`   | `scheme`   | `scheme`   | `protocol` (`"https:"`) |
// | `User`     | `user`     | — | `username` |
// | `Pass`     | `pass`     | — | `password` |
// | `Host`     | `host`     | — | `hostname` |
// | `Port`     | `port`     | — | `port` |
// | `Netloc`   | —          | `netloc` (`"u:p@h:8080"`) | composite |
// | `Path`     | `path`     | `path`     | `pathname` |
// | `Query`    | `query`    | `query`    | `search` (`"?a=1"`) |
// | `Fragment` | `fragment` | `fragment` | `hash` (`"#top"`) |
//
// # The flag that matters: WHATWG normalizes, php and python do NOT
//
// Measured against the real runtimes:
//
// | input | CPython `urlsplit` | WHATWG `URL` |
// |---|---|---|
// | `HTTP://Example.COM/a/../b` | `Example.COM`, `/a/../b` | `example.com`, `/b` |
// | `http://x.com:80/p` | `x.com:80` | `x.com` (default port dropped) |
// | `http://x.com/a b` | `/a b` | `/a%20b` |
//
// CPython `urlsplit` and php `parse_url` are pure SYNTACTIC splits; WHATWG
// folds host case, resolves dot segments, drops default ports and escapes.
// A language declares which it wants via [`ParseMode`] — this is the one place
// the three surfaces genuinely disagree, so it is a flag rather than a fork.

/// A canonical URL component. Every language names a subset of these.
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum UrlField {
    Scheme,
    User,
    Pass,
    Host,
    Port,
    /// `[user[:pass]@]host[:port]` — python's composite, which WHATWG splits.
    Netloc,
    Path,
    Query,
    Fragment,
}

/// Per-language parse rules over ONE parser.
///
/// The languages differ at the GRAMMAR level, not in what a URL is — so these
/// are flags on the shared primitive rather than separate implementations.
/// Measured against the real runtimes on `"HTTP://Example.COM:80/a/../b"`:
///
/// | | scheme | host | port | path |
/// |---|---|---|---|---|
/// | php `parse_url` | `HTTP` | `Example.COM` | `80` | `/a/../b` |
/// | python `urlsplit` | `http` | `Example.COM` | `80` | `/a/../b` |
/// | js WHATWG | `http` | `example.com` | dropped | `/b` |
///
/// python lowercases ONLY the scheme; php preserves everything; WHATWG folds
/// all of it.
#[derive(Clone, Copy)]
pub struct ParseOptions {
    pub mode: ParseMode,
    /// Lowercase the scheme — python `urlsplit` does, php `parse_url` does not.
    pub lowercase_scheme: bool,
}

impl ParseOptions {
    /// python `urlsplit` / `urlparse`.
    pub const fn python() -> ParseOptions {
        ParseOptions {
            mode: ParseMode::Syntactic,
            lowercase_scheme: true,
        }
    }
    /// php `parse_url` — preserves everything it captures.
    pub const fn php() -> ParseOptions {
        ParseOptions {
            mode: ParseMode::Syntactic,
            lowercase_scheme: false,
        }
    }
    /// js `new URL(...)` — WHATWG normalization does the lowering itself.
    pub const fn whatwg() -> ParseOptions {
        ParseOptions {
            mode: ParseMode::Whatwg,
            lowercase_scheme: false,
        }
    }
}

/// Whether a parse normalizes.
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum ParseMode {
    /// php `parse_url`, python `urlsplit` — split only, change nothing.
    Syntactic,
    /// js `new URL(...)` — WHATWG normalization.
    Whatwg,
}

/// The WHATWG property backing a component, and whether a prefix/suffix
/// character has to come off. `None` means the component is composite and is
/// built rather than read.
fn whatwg_property(field: UrlField) -> Option<(&'static str, Trim)> {
    match field {
        UrlField::Scheme => Some(("protocol", Trim::TrailingColon)),
        UrlField::User => Some(("username", Trim::None)),
        UrlField::Pass => Some(("password", Trim::None)),
        UrlField::Host => Some(("hostname", Trim::None)),
        UrlField::Port => Some(("port", Trim::None)),
        UrlField::Path => Some(("pathname", Trim::None)),
        UrlField::Query => Some(("search", Trim::LeadingChar)),
        UrlField::Fragment => Some(("hash", Trim::LeadingChar)),
        UrlField::Netloc => None,
    }
}

enum Trim {
    None,
    /// Drop `?` from `search` / `#` from `hash`. Empty stays empty —
    /// `substring(1)` of `""` is `""`, so no guard is needed.
    LeadingChar,
    /// Drop the `:` that WHATWG's `protocol` carries.
    TrailingColon,
}

/// Read one canonical component from a parsed URL in `url_slot`.
/// Stack: `-> [string]`.
pub fn emit_component(
    chunks: &mut [Chunk],
    current: usize,
    url_slot: u16,
    field: UrlField,
    line: u32,
) {
    let Some((prop, trim)) = whatwg_property(field) else {
        return emit_netloc(chunks, current, url_slot, line);
    };
    lget_at(chunks, current, url_slot, line);
    struct_get(chunks, current, prop, line);
    match trim {
        Trim::None => {}
        Trim::LeadingChar => {
            chunks[current].emit_i32_const(1, line);
            chunks[current].emit_i32_const(0x7FFF_FFFF, line);
            call_import(chunks, current, "ecma:string", "substring", 3, line);
        }
        Trim::TrailingColon => {
            chunks[current].emit_i32_const(0, line);
            chunks[current].emit_i32_const(-1, line);
            call_import(chunks, current, "ecma:string", "slice", 3, line);
        }
    }
}

/// `[user[:pass]@]host[:port]` — WHATWG splits the credentials out, so put them
/// back. `host` already carries `:port`.
fn emit_netloc(chunks: &mut [Chunk], current: usize, url_slot: u16, line: u32) {
    let out = chunks[current].alloc_scratch(1);
    let user = chunks[current].alloc_scratch(1);

    lget_at(chunks, current, url_slot, line);
    struct_get(chunks, current, "username", line);
    lset_at(chunks, current, user, line);
    lget_at(chunks, current, url_slot, line);
    struct_get(chunks, current, "host", line);
    lset_at(chunks, current, out, line);

    lget_at(chunks, current, user, line);
    call_import(chunks, current, "wasm:js-string", "length", 1, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_NE, line);
    chunks[current].emit_if(line);
    {
        let pass = chunks[current].alloc_scratch(1);
        lget_at(chunks, current, url_slot, line);
        struct_get(chunks, current, "password", line);
        lset_at(chunks, current, pass, line);

        lget_at(chunks, current, user, line);
        lget_at(chunks, current, pass, line);
        call_import(chunks, current, "wasm:js-string", "length", 1, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op(Op::I32_NE, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_string_const(":", line);
        lget_at(chunks, current, pass, line);
        call_import(chunks, current, "wasm:js-string", "concat", 2, line);
        chunks[current].emit_else(line);
        chunks[current].emit_string_const("", line);
        chunks[current].emit_end(line);
        call_import(chunks, current, "wasm:js-string", "concat", 2, line);

        chunks[current].emit_string_const("@", line);
        call_import(chunks, current, "wasm:js-string", "concat", 2, line);
        lget_at(chunks, current, out, line);
        call_import(chunks, current, "wasm:js-string", "concat", 2, line);
        lset_at(chunks, current, out, line);
    }
    chunks[current].emit_end(line);
    lget_at(chunks, current, out, line);
}

/// Parse a URL string into the WHATWG object every component read goes through.
/// Stack: `[url]` → `[parsed]`.
pub fn emit_parse(chunks: &mut [Chunk], current: usize, opts: ParseOptions, line: u32) {
    match opts.mode {
        ParseMode::Whatwg => call_import(chunks, current, "web:url", "new", 1, line),
        ParseMode::Syntactic => emit_parse_syntactic(chunks, current, opts, line),
    }
}

fn lget_at(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset_at(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn struct_get(chunks: &mut [Chunk], current: usize, key: &str, line: u32) {
    let k = chunks[current].add_constant(Value::String(Arc::from(key)));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, k, line);
}

/// RFC 3986 §B — the reference regex for splitting a URI, extended with the
/// userinfo and port captures. This is what php's `parse_url` prelude used and
/// what CPython's `urlsplit` implements by hand: a SYNTACTIC split that changes
/// nothing it captures.
///
/// Groups: 1 scheme, 2 user, 3 pass, 4 host, 5 port, 6 path, 7 query, 8 fragment.
const RFC3986_SPLIT: &str = concat!(
    r"^(?:([^:/?#]+):)?",
    r"(?://(?:([^:@/?#]+)(?::([^@/?#]*))?@)?([^:/?#]*)(?::(\d+))?)?",
    r"([^?#]*)(?:\?([^#]*))?(?:#(.*))?$"
);

/// Group index → the WHATWG property name [`emit_component`] reads, and the
/// punctuation WHATWG carries on it.
///
/// **This is what makes the two modes compatible**: a syntactic parse produces
/// exactly the shape a WHATWG parse does, so every component read works against
/// either without knowing which produced it.
const SYNTACTIC_FIELDS: &[(i32, &str, &str, &str)] = &[
    // (group, property, prefix, suffix)
    (1, "protocol", "", ":"),
    (2, "username", "", ""),
    (3, "password", "", ""),
    (4, "hostname", "", ""),
    (5, "port", "", ""),
    (6, "pathname", "", ""),
    (7, "search", "?", ""),
    (8, "hash", "#", ""),
];

/// Split a URL string WITHOUT normalizing. Stack: `[url]` → `[parsed]`.
///
/// Preserves what WHATWG would fold: `HTTP://Example.COM:80/a/../b` keeps its
/// scheme case, host case, default port and dot segments — which is what php
/// `parse_url` and python `urlsplit` both specify.
fn emit_parse_syntactic(chunks: &mut [Chunk], current: usize, opts: ParseOptions, line: u32) {
    let url = chunks[current].alloc_scratch(1);
    let m = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let g = chunks[current].alloc_scratch(1);

    lset_at(chunks, current, url, line);

    chunks[current].emit_string_const(RFC3986_SPLIT, line);
    call_import(chunks, current, "ecma:regexp", "new", 1, line);
    lget_at(chunks, current, url, line);
    call_import(chunks, current, "ecma:regexp", "exec", 2, line);
    lset_at(chunks, current, m, line);

    chunks[current].emit_struct_new(0, 0, line);
    lset_at(chunks, current, out, line);

    for (group, prop, prefix, suffix) in SYNTACTIC_FIELDS {
        // g = m[group] ?? ""   — an unmatched optional group is null/undefined.
        lget_at(chunks, current, m, line);
        chunks[current].emit_i32_const(*group, line);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        lset_at(chunks, current, g, line);
        lget_at(chunks, current, g, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_string_const("", line);
        chunks[current].emit_else(line);
        // Present: re-attach the punctuation WHATWG carries, so the shared
        // component reader's trims apply identically to both modes.
        if !prefix.is_empty() {
            chunks[current].emit_string_const(prefix, line);
        }
        lget_at(chunks, current, g, line);
        if !prefix.is_empty() {
            call_import(chunks, current, "wasm:js-string", "concat", 2, line);
        }
        if !suffix.is_empty() {
            chunks[current].emit_string_const(suffix, line);
            call_import(chunks, current, "wasm:js-string", "concat", 2, line);
        }
        // Grammar-level: python `urlsplit` lowercases the scheme and nothing
        // else; php `parse_url` preserves it. One parser, one flag.
        if *prop == "protocol" && opts.lowercase_scheme {
            call_import(chunks, current, "ecma:string", "toLowerCase", 1, line);
        }
        chunks[current].emit_end(line);

        // out[prop] = value
        let k = chunks[current].add_constant(Value::String(Arc::from(*prop)));
        let v = chunks[current].alloc_scratch(1);
        lset_at(chunks, current, v, line);
        lget_at(chunks, current, out, line);
        lget_at(chunks, current, v, line);
        chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, k, line);
    }

    // `host` is `hostname[:port]` — WHATWG exposes it, and `Netloc` builds on it.
    emit_host_composite(chunks, current, out, line);
    lget_at(chunks, current, out, line);
}

/// `out.host = hostname + (port ? ":" + port : "")`.
fn emit_host_composite(chunks: &mut [Chunk], current: usize, out: u16, line: u32) {
    let port = chunks[current].alloc_scratch(1);
    lget_at(chunks, current, out, line);
    struct_get(chunks, current, "port", line);
    lset_at(chunks, current, port, line);

    lget_at(chunks, current, out, line);
    lget_at(chunks, current, out, line);
    struct_get(chunks, current, "hostname", line);
    lget_at(chunks, current, port, line);
    call_import(chunks, current, "wasm:js-string", "length", 1, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_NE, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const(":", line);
    lget_at(chunks, current, port, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_end(line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    let k = chunks[current].add_constant(Value::String(Arc::from("host")));
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, k, line);
}

/// Read one canonical component from a parsed URL **on the stack**.
/// Stack: `[parsed]` → `[string]`.
///
/// The slot-shaped [`emit_component`] suits an emitter that already stashed the
/// parsed URL; a PROFILE builtin instead receives its argument on the stack, so
/// java's `__j_url_*` getters and any other profile-bound reader use this form.
/// Both go through the same reshaping — there is one component model.
pub fn emit_component_of(chunks: &mut [Chunk], current: usize, field: UrlField, line: u32) {
    let slot = chunks[current].alloc_scratch(1);
    lset_at(chunks, current, slot, line);
    emit_component(chunks, current, slot, field, line);
}
