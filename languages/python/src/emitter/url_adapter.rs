//! Python `urllib.parse` over the WHATWG URL surface.
//!
//! `web:url` is already the URL parser for JS (`new URL(...)`,
//! `URLSearchParams`) and node's `url` module, so nothing new is registered
//! here.
//!
//! **The WHATWG→canonical reshaping lives in `primitives::url`**, not here —
//! `protocol` minus its trailing colon, the composite `netloc`, the `?`/`#`
//! prefixes. php `parse_url` needs exactly the same reshaping, so it is shared
//! (`url::UrlField`, `url::emit_component`). What remains in this file is only
//! what is python's: the `SplitResult` named tuple, `parse_qs`/`parse_qsl`
//! shapes, and python's argument conventions.
//!
//! `urljoin` needs no reshaping at all: WHATWG relative-URL resolution against
//! a base IS RFC 3986 §5, which is what CPython implements.
//!
//! Arguments arrive pre-pushed on the stack, left to right — the `emit_common`
//! convention shared with `os_path_adapter` / `math_adapter`.

use vybe_runtime::Chunk;
use vybe_compiler::primitives::url;
use vybe_runtime::opcode::Op;

use vybe_compiler::primitives::tuples;

/// `substring(start, END)` runs to the end of the string.
const END: i32 = 0x7FFF_FFFF;

fn call_import(chunks: &mut [Chunk], current: usize, module: &str, name: &str, argc: u8, line: u32) {
    // Register on the CURRENT chunk — an index from chunks[0] resolves to the
    // wrong host fn when the code runs in a function chunk.
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

fn stash_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(argc as u16);
    for offset in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    base
}

// ── percent-encoding ──────────────────────────────────────────────────
//
// The CODEC is `primitives::url`; only python's argument handling is here.

fn emit_percent(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
    opts: url::PercentOptions,
    decode: bool,
) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    lget(chunks, current, base, line);
    if decode {
        url::emit_percent_decode(chunks, current, opts, line);
    } else {
        url::emit_percent_encode(chunks, current, opts, line);
    }
}

/// `quote(s)` — CPython's default `safe='/'`, so `/` survives.
pub fn emit_quote(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_percent(chunks, current, argc, line, url::PercentOptions::path(), false);
}

/// `quote_plus(s)` — space becomes `+` AND `/` is escaped.
///
/// Deliberately not `path()` plus a space fixup: CPython escapes `/` here.
/// Building this on `quote` (which keeps `/` safe) made `quote_plus("a b/c")`
/// return `a+b/c` where python gives `a+b%2Fc`.
pub fn emit_quote_plus(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_percent(chunks, current, argc, line, url::PercentOptions::form_rfc3986(), false);
}

/// `unquote(s)` — `+` stays literal.
pub fn emit_unquote(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_percent(chunks, current, argc, line, url::PercentOptions::rfc3986(), true);
}

/// `unquote_plus(s)` — `+` decodes back to a space.
pub fn emit_unquote_plus(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_percent(chunks, current, argc, line, url::PercentOptions::form_rfc3986(), true);
}

// ── structural: split / unsplit / join ────────────────────────────────
//
// `web:url.new` is the WHATWG parser already registered for JS `new URL(...)`;
// the reshapers above turn its field names into CPython's.

/// `urlsplit(url)` / `urlparse(url)` → `SplitResult(scheme, netloc, path,
/// query, fragment)`, a NAMED tuple so `.scheme` etc. read by name.
pub fn emit_urlsplit(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    let parsed = chunks[current].alloc_scratch(1);

    lget(chunks, current, base, line);
    url::emit_parse(chunks, current, url::ParseOptions::python(), line);
    lset(chunks, current, parsed, line);

    // Component reads are the SHARED superset — the WHATWG→canonical reshaping
    // (`protocol` minus its colon, the composite `netloc`, `?`/`#` prefixes)
    // is identical for php `parse_url` and belongs in one place.
    for field in [
        url::UrlField::Scheme,
        url::UrlField::Netloc,
        url::UrlField::Path,
        url::UrlField::Query,
        url::UrlField::Fragment,
    ] {
        url::emit_component(chunks, current, parsed, field, line);
    }

    chunks[current].emit_array_new_fixed(0, 5, line);
    let fields = [
        Some("scheme".to_string()),
        Some("netloc".to_string()),
        Some("path".to_string()),
        Some("query".to_string()),
        Some("fragment".to_string()),
    ];
    tuples::emit_named_tuple(chunks, current, &fields, Some("SplitResult"), line);
}

/// `urlunsplit(parts)` / `urlunparse(parts)` — reassemble, and it must
/// round-trip: `urlunsplit(urlsplit(u)) == u`.
///
/// Each separator is emitted only when its component is non-empty, which is
/// what makes a URL with no query or no fragment come back unchanged.
pub fn emit_urlunsplit(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    let out = chunks[current].alloc_scratch(1);
    let part = chunks[current].alloc_scratch(1);

    // scheme -> "scheme:" when present
    emit_part(chunks, current, base, 0, part, line);
    emit_if_non_empty(chunks, current, part, line);
    lget(chunks, current, part, line);
    chunks[current].emit_string_const(":", line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_end(line);
    lset(chunks, current, out, line);

    emit_append_prefixed(chunks, current, base, 1, out, part, "//", line);
    emit_append_prefixed(chunks, current, base, 2, out, part, "", line);
    emit_append_prefixed(chunks, current, base, 3, out, part, "?", line);
    emit_append_prefixed(chunks, current, base, 4, out, part, "#", line);

    lget(chunks, current, out, line);
}

/// `parts[i]` into `slot`.
fn emit_part(chunks: &mut [Chunk], current: usize, base: u16, i: i32, slot: u16, line: u32) {
    lget(chunks, current, base, line);
    chunks[current].emit_i32_const(i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    lset(chunks, current, slot, line);
}

/// Push i32 bool `slot.length != 0` and open an `if`.
fn emit_if_non_empty(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    lget(chunks, current, slot, line);
    call_import(chunks, current, "wasm:js-string", "length", 1, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_NE, line);
    chunks[current].emit_if_value(line);
}

/// `out += prefix + parts[i]`, but only when `parts[i]` is non-empty.
fn emit_append_prefixed(
    chunks: &mut [Chunk],
    current: usize,
    base: u16,
    i: i32,
    out: u16,
    part: u16,
    prefix: &str,
    line: u32,
) {
    emit_part(chunks, current, base, i, part, line);
    lget(chunks, current, out, line);
    emit_if_non_empty(chunks, current, part, line);
    if prefix.is_empty() {
        lget(chunks, current, part, line);
    } else {
        chunks[current].emit_string_const(prefix, line);
        lget(chunks, current, part, line);
        call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_end(line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    lset(chunks, current, out, line);
}

/// `urljoin(base, url)` — WHATWG relative resolution against a base IS
/// RFC 3986 §5, which is what CPython implements, so no reshaping is needed.
pub fn emit_urljoin(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    lget(chunks, current, base + 1, line);
    lget(chunks, current, base, line);
    call_import(chunks, current, "web:url", "new", 2, line);
    call_import(chunks, current, "web:url", "urlToString", 1, line);
}

// ── query strings ─────────────────────────────────────────────────────
//
// `web:url`'s `searchParams*` surface is the WHATWG query parser, already
// registered for JS `URLSearchParams`. CPython's three shapes are views over
// the same pairs: `parse_qsl` the ordered pair list, `parse_qs` those pairs
// grouped by key, `urlencode` the inverse.

/// Build a `URLSearchParams` over the query string at `slot`.
fn search_params(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    lget(chunks, current, slot, line);
    call_import(chunks, current, "web:url", "searchParamsNew", 1, line);
}

/// `parse_qsl(q)` → `[(key, value), …]` in order, duplicates preserved.
pub fn emit_parse_qsl(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_array_new_fixed(0, 0, line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    search_params(chunks, current, base, line);
    call_import(chunks, current, "web:url", "searchParamsEntries", 1, line);
    // `searchParamsEntries` yields 2-element arrays; CPython yields TUPLES,
    // and the difference is visible in `repr` — `('a', '1')` not `['a', '1']`.
    emit_tag_pairs_as_tuples(chunks, current, line);
}

/// `parse_qs(q)` → `{key: [value, …]}`, grouping duplicates.
pub fn emit_parse_qs(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        call_import(chunks, current, "ecma:map", "new", 0, line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    let params = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let keys = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);

    search_params(chunks, current, base, line);
    lset(chunks, current, params, line);
    call_import(chunks, current, "ecma:map", "new", 0, line);
    lset(chunks, current, out, line);

    // Distinct keys, in first-seen order.
    lget(chunks, current, params, line);
    call_import(chunks, current, "web:url", "searchParamsKeys", 1, line);
    lset(chunks, current, keys, line);
    chunks[current].emit_i32_const(0, line);
    lset(chunks, current, i, line);
    lget(chunks, current, keys, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    lset(chunks, current, n, line);

    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    lget(chunks, current, i, line);
    lget(chunks, current, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    lget(chunks, current, keys, line);
    lget(chunks, current, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    lset(chunks, current, key, line);

    // out[key] = params.getAll(key)  — idempotent for a repeated key.
    lget(chunks, current, out, line);
    lget(chunks, current, key, line);
    lget(chunks, current, params, line);
    lget(chunks, current, key, line);
    call_import(chunks, current, "web:url", "searchParamsGetAll", 2, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    lget(chunks, current, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(chunks, current, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    lget(chunks, current, out, line);
}

/// `urlencode(mapping)` → `k=v&k=v`, each part `quote_plus`d.
///
/// CPython uses `quote_plus` here, so a space is `+` and `&` is `%26` —
/// `{"name": "Alice & Bob"}` gives `name=Alice+%26+Bob`.
pub fn emit_urlencode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    let params = chunks[current].alloc_scratch(1);

    call_import(chunks, current, "web:url", "searchParamsNew", 0, line);
    lset(chunks, current, params, line);

    let entries = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let pair = chunks[current].alloc_scratch(1);

    lget(chunks, current, base, line);
    call_import(chunks, current, "ecma:object", "entries", 1, line);
    lset(chunks, current, entries, line);
    chunks[current].emit_i32_const(0, line);
    lset(chunks, current, i, line);
    lget(chunks, current, entries, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    lset(chunks, current, n, line);

    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    lget(chunks, current, i, line);
    lget(chunks, current, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    lget(chunks, current, entries, line);
    lget(chunks, current, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    lset(chunks, current, pair, line);

    lget(chunks, current, params, line);
    lget(chunks, current, pair, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    lget(chunks, current, pair, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    call_import(chunks, current, "web:url", "searchParamsAppend", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    lget(chunks, current, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(chunks, current, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    lget(chunks, current, params, line);
    call_import(chunks, current, "web:url", "searchParamsToString", 1, line);
}

/// Re-tag each `[k, v]` pair in the array on TOS as a TUPLE, so `repr` prints
/// `('a', '1')` the way CPython's `parse_qsl` does.
fn emit_tag_pairs_as_tuples(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);

    lset(chunks, current, arr, line);
    chunks[current].emit_i32_const(0, line);
    lset(chunks, current, i, line);
    lget(chunks, current, arr, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    lset(chunks, current, n, line);

    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    lget(chunks, current, i, line);
    lget(chunks, current, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    lget(chunks, current, arr, line);
    lget(chunks, current, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    tuples::emit_tag(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    lget(chunks, current, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(chunks, current, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    // Only the PAIRS are tuples. `parse_qsl` returns a LIST of them —
    // tagging the outer array too printed `(('a','1'), …)` for `[('a','1'), …]`.
    lget(chunks, current, arr, line);
}
