//! Python `urllib.parse` over the WHATWG URL surface.
//!
//! `web:url` is already the URL parser for JS (`new URL(...)`,
//! `URLSearchParams`) and node's `url` module, so nothing new is registered
//! here — this only reshapes WHATWG's field names into CPython's.
//!
//! The two spellings differ in punctuation and in one composite field:
//!
//! | CPython `SplitResult` | WHATWG URL |
//! |---|---|
//! | `scheme`   (`"https"`)              | `protocol` (`"https:"`) |
//! | `netloc`   (`"u:p@host:8080"`)      | `username`+`password`+`host` |
//! | `path`                              | `pathname` |
//! | `query`    (`"a=1"`)                | `search`   (`"?a=1"`) |
//! | `fragment` (`"top"`)                | `hash`     (`"#top"`) |
//!
//! `urljoin` needs no reshaping at all: WHATWG relative-URL resolution against
//! a base IS RFC 3986 §5, which is what CPython implements.
//!
//! Arguments arrive pre-pushed on the stack, left to right — the `emit_common`
//! convention shared with `os_path_adapter` / `math_adapter`.

use vybe_runtime::Chunk;
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

fn sget(chunks: &mut [Chunk], current: usize, key: &str, line: u32) {
    let k = chunks[current].add_constant(vybe_runtime::Value::String(std::sync::Arc::from(key)));
    chunks[current].emit_op_u16(Op::STRUCT_GET, k, line);
}

/// Pop `argc` values (deepest first) into fresh scratch slots; `base + i` is
/// the i-th argument.
fn stash_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(argc as u16);
    for offset in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    base
}

/// Read `url[field]` and drop its leading punctuation character (`?` from
/// `search`, `#` from `hash`). An empty field stays empty — `substring(1)` on
/// `""` is `""`, so no guard is needed.
fn field_without_prefix(chunks: &mut [Chunk], current: usize, url: u16, field: &str, line: u32) {
    lget(chunks, current, url, line);
    sget(chunks, current, field, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_i32_const(END, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
}

/// Read `url.protocol` (`"https:"`) as CPython's `scheme` (`"https"`) by
/// dropping the trailing colon.
fn scheme_field(chunks: &mut [Chunk], current: usize, url: u16, line: u32) {
    lget(chunks, current, url, line);
    sget(chunks, current, "protocol", line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_i32_const(-1, line);
    call_import(chunks, current, "ecma:string", "slice", 3, line);
}

/// CPython's `netloc` — `[user[:password]@]host[:port]`. WHATWG splits the
/// credentials out, so recombine them; `host` already carries `:port`.
fn netloc_field(chunks: &mut [Chunk], current: usize, url: u16, line: u32) {
    let out = chunks[current].alloc_scratch(1);
    let user = chunks[current].alloc_scratch(1);

    lget(chunks, current, url, line);
    sget(chunks, current, "username", line);
    lset(chunks, current, user, line);

    lget(chunks, current, url, line);
    sget(chunks, current, "host", line);
    lset(chunks, current, out, line);

    // username non-empty → prepend "user[:pass]@"
    lget(chunks, current, user, line);
    call_import(chunks, current, "wasm:js-string", "length", 1, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_NE, line);
    chunks[current].emit_if(line);
    {
        let pass = chunks[current].alloc_scratch(1);
        lget(chunks, current, url, line);
        sget(chunks, current, "password", line);
        lset(chunks, current, pass, line);

        lget(chunks, current, user, line);
        // password non-empty → ":pass"
        lget(chunks, current, pass, line);
        call_import(chunks, current, "wasm:js-string", "length", 1, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op(Op::I32_NE, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_string_const(":", line);
        lget(chunks, current, pass, line);
        call_import(chunks, current, "wasm:js-string", "concat", 2, line);
        chunks[current].emit_else(line);
        chunks[current].emit_string_const("", line);
        chunks[current].emit_end(line);
        call_import(chunks, current, "wasm:js-string", "concat", 2, line);

        chunks[current].emit_string_const("@", line);
        call_import(chunks, current, "wasm:js-string", "concat", 2, line);
        lget(chunks, current, out, line);
        call_import(chunks, current, "wasm:js-string", "concat", 2, line);
        lset(chunks, current, out, line);
    }
    chunks[current].emit_end(line);

    lget(chunks, current, out, line);
}

/// `urljoin(base, url)` → WHATWG relative resolution, which IS RFC 3986 §5.
/// Stack: `[base, url] -> [string]`.
pub fn emit_urljoin(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        chunks[current].emit_string_const("", line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    let rel = base + 1;
    // web:url.new(input, base) — note the argument ORDER is the reverse of
    // CPython's, which takes the base first.
    lget(chunks, current, rel, line);
    lget(chunks, current, base, line);
    call_import(chunks, current, "web:url", "new", 2, line);
    let url = chunks[current].alloc_scratch(1);
    lset(chunks, current, url, line);

    // Unparseable → CPython returns the relative reference unchanged.
    lget(chunks, current, url, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    lget(chunks, current, rel, line);
    chunks[current].emit_else(line);
    lget(chunks, current, url, line);
    sget(chunks, current, "href", line);
    chunks[current].emit_end(line);
}

/// `urlsplit(url)` → `SplitResult(scheme, netloc, path, query, fragment)`.
///
/// A NAMED tuple, so `.scheme`/`.netloc`/… resolve through the shared
/// named-tuple metadata (`tuples::emit_named_tuple` stamps a by-name key per
/// field) while `urlunsplit(split)` still receives the plain 5-element array
/// it indexes. Both forms come free from one shape; nothing needs a walker
/// rewrite of the attribute names.
/// Stack: `[url] -> [tuple]`.
pub fn emit_urlsplit(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_op(Op::NULL, line);
    }
    let arg = chunks[current].alloc_scratch(1);
    lset(chunks, current, arg, line);

    lget(chunks, current, arg, line);
    call_import(chunks, current, "web:url", "new", 1, line);
    let url = chunks[current].alloc_scratch(1);
    lset(chunks, current, url, line);

    // A relative reference has no scheme, so WHATWG refuses it outright. Give
    // back CPython's shape for that case: everything in `path`.
    lget(chunks, current, url, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    {
        chunks[current].emit_string_const("", line);
        chunks[current].emit_string_const("", line);
        lget(chunks, current, arg, line);
        chunks[current].emit_string_const("", line);
        chunks[current].emit_string_const("", line);
        emit_split_result(chunks, current, line);
    }
    chunks[current].emit_else(line);
    {
        scheme_field(chunks, current, url, line);
        netloc_field(chunks, current, url, line);
        lget(chunks, current, url, line);
        sget(chunks, current, "pathname", line);
        field_without_prefix(chunks, current, url, "search", line);
        field_without_prefix(chunks, current, url, "hash", line);
        emit_split_result(chunks, current, line);
    }
    chunks[current].emit_end(line);
}

/// Pack the five components on the stack into CPython's `SplitResult`.
/// Stack: `[scheme, netloc, path, query, fragment] -> [named tuple]`.
fn emit_split_result(chunks: &mut [Chunk], current: usize, line: u32) {
    const FIELDS: [&str; 5] = ["scheme", "netloc", "path", "query", "fragment"];
    let base = chunks[current].alloc_scratch(5);
    vybe_compiler::primitives::collections::emit_pack_n(chunks, current, 5, base, line);
    let names: Vec<Option<String>> = FIELDS.iter().map(|f| Some((*f).to_string())).collect();
    tuples::emit_named_tuple(chunks, current, &names, Some("SplitResult"), line);
}

/// `urlunsplit((scheme, netloc, path, query, fragment))` — the textual inverse
/// of `urlsplit`, per RFC 3986 §5.3: each component contributes its own
/// delimiter and an empty component contributes nothing.
/// Stack: `[tuple] -> [string]`.
pub fn emit_urlunsplit(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
        return;
    }
    let parts = chunks[current].alloc_scratch(1);
    lset(chunks, current, parts, line);
    let out = chunks[current].alloc_scratch(1);
    let piece = chunks[current].alloc_scratch(1);

    chunks[current].emit_string_const("", line);
    lset(chunks, current, out, line);

    // (index, prefix-when-non-empty)
    for (index, prefix) in [(0usize, ":"), (1, "//"), (2, ""), (3, "?"), (4, "#")] {
        lget(chunks, current, parts, line);
        chunks[current].emit_i32_const(index as i32, line);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        lset(chunks, current, piece, line);

        lget(chunks, current, piece, line);
        call_import(chunks, current, "wasm:js-string", "length", 1, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op(Op::I32_NE, line);
        chunks[current].emit_if(line);
        {
            lget(chunks, current, out, line);
            // `scheme` takes its delimiter AFTER, every other one BEFORE.
            if index == 0 {
                lget(chunks, current, piece, line);
                call_import(chunks, current, "wasm:js-string", "concat", 2, line);
                chunks[current].emit_string_const(prefix, line);
                call_import(chunks, current, "wasm:js-string", "concat", 2, line);
            } else {
                chunks[current].emit_string_const(prefix, line);
                call_import(chunks, current, "wasm:js-string", "concat", 2, line);
                lget(chunks, current, piece, line);
                call_import(chunks, current, "wasm:js-string", "concat", 2, line);
            }
            lset(chunks, current, out, line);
        }
        chunks[current].emit_end(line);
    }

    lget(chunks, current, out, line);
}

/// `urlencode(dict)` → `"a=1&b=2"` with `application/x-www-form-urlencoded`
/// escaping, which is what `URLSearchParams.toString()` produces — including
/// the `+` for a space that `quote` does NOT use.
/// Stack: `[dict] -> [string]`.
pub fn emit_urlencode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
        return;
    }
    let dict = chunks[current].alloc_scratch(1);
    lset(chunks, current, dict, line);

    let params = chunks[current].alloc_scratch(1);
    call_import(chunks, current, "web:url", "searchParamsNew", 0, line);
    lset(chunks, current, params, line);

    // A Python dict is a Map; iterate its entries as [key, value] pairs.
    let entries = chunks[current].alloc_scratch(1);
    lget(chunks, current, dict, line);
    call_import(chunks, current, "ecma:map", "entries", 1, line);
    call_import(chunks, current, "ecma:array", "from", 1, line);
    lset(chunks, current, entries, line);

    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let pair = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    lset(chunks, current, i, line);
    lget(chunks, current, entries, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    lset(chunks, current, n, line);

    let block = chunks[current].emit_block(line);
    let (lp, _) = chunks[current].emit_loop_s(line);
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

/// `parse_qs(query)` → `{key: [values]}`. CPython gives every key a LIST,
/// which is what distinguishes it from `parse_qsl`.
/// Stack: `[query] -> [dict]`.
pub fn emit_parse_qs(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
    }
    let query = chunks[current].alloc_scratch(1);
    lset(chunks, current, query, line);

    let params = chunks[current].alloc_scratch(1);
    lget(chunks, current, query, line);
    call_import(chunks, current, "web:url", "searchParamsNew", 1, line);
    lset(chunks, current, params, line);

    let out = chunks[current].alloc_scratch(1);
    call_import(chunks, current, "ecma:map", "new", 0, line);
    lset(chunks, current, out, line);

    let keys = chunks[current].alloc_scratch(1);
    lget(chunks, current, params, line);
    call_import(chunks, current, "web:url", "searchParamsKeys", 1, line);
    call_import(chunks, current, "ecma:array", "from", 1, line);
    lset(chunks, current, keys, line);

    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    lset(chunks, current, i, line);
    lget(chunks, current, keys, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    lset(chunks, current, n, line);

    let block = chunks[current].emit_block(line);
    let (lp, _) = chunks[current].emit_loop_s(line);
    lget(chunks, current, i, line);
    lget(chunks, current, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    lget(chunks, current, keys, line);
    lget(chunks, current, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    lset(chunks, current, key, line);

    // out.set(key, [...getAll(key)]) — repeated keys collapse to one entry
    // because `set` overwrites, matching CPython's single list per key.
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

/// `parse_qsl(query)` → `[(key, value), …]` — the ORDERED pair list, keeping
/// duplicate keys as separate entries (the difference from `parse_qs`, which
/// groups them into one list per key).
/// Stack: `[query] -> [list of tuples]`.
pub fn emit_parse_qsl(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
    }
    let query = chunks[current].alloc_scratch(1);
    lset(chunks, current, query, line);

    let entries = chunks[current].alloc_scratch(1);
    lget(chunks, current, query, line);
    call_import(chunks, current, "web:url", "searchParamsNew", 1, line);
    call_import(chunks, current, "web:url", "searchParamsEntries", 1, line);
    call_import(chunks, current, "ecma:array", "from", 1, line);
    lset(chunks, current, entries, line);

    // The entries arrive as plain 2-element arrays; Python wants TUPLES, which
    // repr as `('a', '1')` rather than `['a', '1']`. Re-tag in place.
    let out = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let pair = chunks[current].alloc_scratch(1);

    let new_len = chunks[current].add_import("vybe:js-array", "newWithLength");
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, new_len, line);
    chunks[current].emit(1u8, line);
    lset(chunks, current, out, line);

    chunks[current].emit_i32_const(0, line);
    lset(chunks, current, i, line);
    lget(chunks, current, entries, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    lset(chunks, current, n, line);

    let block = chunks[current].emit_block(line);
    let (lp, _) = chunks[current].emit_loop_s(line);
    lget(chunks, current, i, line);
    lget(chunks, current, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    lget(chunks, current, entries, line);
    lget(chunks, current, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    lset(chunks, current, pair, line);

    lget(chunks, current, out, line);
    lget(chunks, current, pair, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    lget(chunks, current, pair, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    tuples::emit_tuple(chunks, current, 2, line);
    call_import(chunks, current, "ecma:array", "push", 2, line);
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

/// `quote(s)` — percent-encoding that leaves `/` alone (CPython's default
/// `safe='/'`). `encodeURIComponent` escapes `/`, so put it back; the reverse
/// swap on a placeholder is not needed because `encodeURIComponent` never
/// produces a bare `/`.
/// Stack: `[s] -> [string]`.
pub fn emit_quote(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
        return;
    }
    // Extra args (`safe=`, `encoding=`) are not modelled; drop them.
    let base = stash_args(chunks, current, argc, line);
    lget(chunks, current, base, line);
    call_import(chunks, current, "ecma:string", "encodeURIComponent", 1, line);
    chunks[current].emit_string_const("%2F", line);
    chunks[current].emit_string_const("/", line);
    call_import(chunks, current, "ecma:string", "replaceAll", 3, line);
}

/// `quote_plus(s)` — `quote` but a space becomes `+`, not `%20`.
/// Stack: `[s] -> [string]`.
pub fn emit_quote_plus(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_quote(chunks, current, argc, line);
    chunks[current].emit_string_const("%20", line);
    chunks[current].emit_string_const("+", line);
    call_import(chunks, current, "ecma:string", "replaceAll", 3, line);
}

/// `unquote(s)`. Stack: `[s] -> [string]`.
pub fn emit_unquote(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    lget(chunks, current, base, line);
    call_import(chunks, current, "ecma:string", "decodeURIComponent", 1, line);
}

/// `unquote_plus(s)` — `+` decodes back to a space before percent-decoding.
/// Stack: `[s] -> [string]`.
pub fn emit_unquote_plus(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    lget(chunks, current, base, line);
    chunks[current].emit_string_const("+", line);
    chunks[current].emit_string_const(" ", line);
    call_import(chunks, current, "ecma:string", "replaceAll", 3, line);
    call_import(chunks, current, "ecma:string", "decodeURIComponent", 1, line);
}
