//! Request body parsing — `application/x-www-form-urlencoded` and
//! `multipart/form-data`.
//!
//! Neither is in any spec host surface. `wasi:http` carries the body as bytes;
//! `node:querystring` covers urlencoded only, and **multipart exists nowhere**.
//! Per `documentation/httpserver.md` §4a this is therefore a PRIMITIVE, and
//! deliberately ONE of them: multipart parsing has a real CVE history, so the
//! project ships a single implementation rather than one per language.
//!
//! One implementation serves PHP `$_POST`/`$_FILES`, Python
//! `cgi.FieldStorage`, Rack's multipart, ASP.NET `IFormFile` and JS
//! `multipart/form-data`.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// Parse `a=1&b=2` into a map. Stack: [string] → [map].
///
/// Percent-decoding and `+`-as-space are the urlencoded rules (RFC 1866
/// §8.2.1), applied through `ecma:string.decodeURIComponent` rather than
/// hand-rolled.
pub fn emit_parse_urlencoded(chunks: &mut [Chunk], current: usize, line: u32) {
    let src = chunks[current].alloc_scratch(1);
    let parts = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let pair = chunks[current].alloc_scratch(1);
    let eq = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, src, line);
    super::collections::emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    // A null or empty body is an empty map, not an error.
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, src, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    chunks[current].emit_string_const("&", line);
    let split = chunks[current].add_import("ecma:string", "split");
    chunks[current].emit_call(split, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, parts, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, parts, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, parts, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pair, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, pair, line);
    chunks[current].emit_string_const("=", line);
    let index_of = chunks[current].add_import("ecma:string", "indexOf");
    chunks[current].emit_call(index_of, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, eq, line);

    // A field with no `=` is skipped; an empty segment (`a=1&&b=2`) too.
    chunks[current].emit_op_u16(Op::LOCAL_GET, eq, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    emit_decode_slice(chunks, current, pair, 0, Some(eq), line);
    emit_decode_slice(chunks, current, pair, 1, None, line);
    super::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// Push `pair`'s name (`part` 0, up to `eq`) or value (`part` 1, after `eq`),
/// `+`-decoded and percent-decoded. Stack: [] → [string].
fn emit_decode_slice(
    chunks: &mut [Chunk],
    current: usize,
    pair: u16,
    part: u8,
    eq: Option<u16>,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, pair, line);
    if part == 0 {
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, eq.expect("name slice needs eq"), line);
    } else {
        // The `=` index is recomputed here so the caller need not keep it live
        // across the name push.
        chunks[current].emit_op_u16(Op::LOCAL_GET, pair, line);
        chunks[current].emit_string_const("=", line);
        let index_of = chunks[current].add_import("ecma:string", "indexOf");
        chunks[current].emit_call(index_of, 2, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, pair, line);
        let len = chunks[current].add_import("wasm:js-string", "length");
        chunks[current].emit_call(len, 1, line);
    }
    let substring = chunks[current].add_import("wasm:js-string", "substring");
    chunks[current].emit_call(substring, 3, line);

    // `+` means space before percent-decoding (RFC 1866 §8.2.1).
    chunks[current].emit_string_const("+", line);
    chunks[current].emit_string_const(" ", line);
    let replace_all = chunks[current].add_import("ecma:string", "replaceAll");
    chunks[current].emit_call(replace_all, 3, line);

    let decode = chunks[current].add_import("ecma:string", "decodeURIComponent");
    chunks[current].emit_call(decode, 1, line);
}

/// Parse a `multipart/form-data` body. Stack: [body, content_type] → [map].
///
/// Returns `{ "fields": map, "files": map }` — plain fields and file uploads
/// kept apart, because every language surfaces them separately (PHP
/// `$_POST`/`$_FILES`, Python `FieldStorage` items with/without `filename`,
/// ASP.NET `Form`/`Files`).
///
/// The body arrives as bytes and is handled as a LATIN-1 string: every byte
/// maps to exactly one char in 0..=255, so the round trip is lossless even for
/// binary uploads, and the delimiter scanning becomes ordinary string work.
/// Each file's content is handed back as that same byte-per-char string; the
/// caller re-encodes when writing it out.
///
/// RFC 7578: parts are separated by `--boundary`, each part is headers, a blank
/// line, then content, and the epilogue begins at `--boundary--`.
pub fn emit_parse_multipart(chunks: &mut [Chunk], current: usize, line: u32) {
    let content_type = chunks[current].alloc_scratch(1);
    let body = chunks[current].alloc_scratch(1);
    let boundary = chunks[current].alloc_scratch(1);
    let parts = chunks[current].alloc_scratch(1);
    let fields = chunks[current].alloc_scratch(1);
    let files = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let part = chunks[current].alloc_scratch(1);
    let split_at = chunks[current].alloc_scratch(1);
    let head = chunks[current].alloc_scratch(1);
    let name = chunks[current].alloc_scratch(1);
    let filename = chunks[current].alloc_scratch(1);
    let upload = chunks[current].alloc_scratch(1);
    let entry = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, content_type, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, body, line);

    super::collections::emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fields, line);
    super::collections::emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, files, line);

    // boundary=... from the Content-Type header.
    emit_after_marker(chunks, current, content_type, "boundary=", line);
    emit_strip_quotes(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, boundary, line);

    // Split on the delimiter. Leading/trailing segments are the preamble and
    // the `--` epilogue; both fail the disposition check below and are skipped.
    chunks[current].emit_op_u16(Op::LOCAL_GET, body, line);
    chunks[current].emit_string_const("--", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, boundary, line);
    super::strings::emit_concat(&mut chunks[current], 2, line);
    let split = chunks[current].add_import("ecma:string", "split");
    chunks[current].emit_call(split, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, parts, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, parts, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, parts, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, part, line);

    // Headers end at the first blank line (CRLFCRLF).
    chunks[current].emit_op_u16(Op::LOCAL_GET, part, line);
    chunks[current].emit_string_const("\r\n\r\n", line);
    let index_of = chunks[current].add_import("ecma:string", "indexOf");
    chunks[current].emit_call(index_of, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, split_at, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, split_at, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, part, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, split_at, line);
    let substring = chunks[current].add_import("wasm:js-string", "substring");
    chunks[current].emit_call(substring, 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, head, line);

    emit_after_marker(chunks, current, head, "name=", line);
    emit_strip_quotes(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, name, line);
    emit_after_marker(chunks, current, head, "filename=", line);
    emit_strip_quotes(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, filename, line);

    // Content runs from after the blank line to the trailing CRLF the next
    // delimiter is prefixed with.
    chunks[current].emit_op_u16(Op::LOCAL_GET, name, line);
    let len_fn = chunks[current].add_import("wasm:js-string", "length");
    chunks[current].emit_call(len_fn, 1, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);

    // A part with a filename is an upload even when the content is empty.
    chunks[current].emit_op_u16(Op::LOCAL_GET, filename, line);
    let len_fn2 = chunks[current].add_import("wasm:js-string", "length");
    chunks[current].emit_call(len_fn2, 1, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    // An upload is not just bytes: the client's filename, the part's declared
    // media type and the byte count are all data the caller needs, and none of
    // them can be recovered from the content afterwards. PHP spells these
    // `name`/`type`/`size` inside `$_FILES[k]`, Rack `:filename`/`:type` — the
    // renaming is each adapter's, so the keys here are the neutral ones.
    chunks[current].emit_op_u16(Op::LOCAL_GET, files, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name, line);

    emit_part_content(chunks, current, part, split_at, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, upload, line);

    super::collections::emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, entry, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, entry, line);
    chunks[current].emit_string_const("filename", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, filename, line);
    super::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, entry, line);
    chunks[current].emit_string_const("type", line);
    emit_after_marker(chunks, current, head, "Content-Type: ", line);
    super::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    // The body is carried as LATIN-1 (one char per byte), so the string length
    // IS the octet count — a UTF-8 length would under-report every binary
    // upload.
    chunks[current].emit_op_u16(Op::LOCAL_GET, entry, line);
    chunks[current].emit_string_const("size", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, upload, line);
    let size_fn = chunks[current].add_import("wasm:js-string", "length");
    chunks[current].emit_call(size_fn, 1, line);
    super::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, entry, line);
    chunks[current].emit_string_const("content", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, upload, line);
    super::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, entry, line);
    super::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fields, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name, line);
    emit_part_content(chunks, current, part, split_at, line);
    super::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    super::collections::emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_string_const("fields", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fields, line);
    super::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_string_const("files", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, files, line);
    super::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// The part's body: after the blank line, minus the trailing CRLF.
/// Stack: [] → [string].
fn emit_part_content(chunks: &mut [Chunk], current: usize, part: u16, split_at: u16, line: u32) {
    let start = chunks[current].alloc_scratch(1);
    let end = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, split_at, line);
    chunks[current].emit_i32_const(4, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, part, line);
    let len = chunks[current].add_import("wasm:js-string", "length");
    chunks[current].emit_call(len, 1, line);
    chunks[current].emit_i32_const(2, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, end, line);

    // Guard a malformed part where the trailing CRLF is missing.
    chunks[current].emit_op_u16(Op::LOCAL_GET, end, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, end, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, part, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, end, line);
    let substring = chunks[current].add_import("wasm:js-string", "substring");
    chunks[current].emit_call(substring, 3, line);
}

/// The text after parameter `marker` up to the next `;`, `\r` or end.
/// Stack: [] → [string]. Empty when the marker is absent.
///
/// The match must begin a parameter — at the start of the string, or after a
/// `;`/space. A bare substring search would find `name=` inside `filename=`,
/// so `filename="x.txt"; name="f"` (RFC 7578 fixes no parameter order, and
/// real clients emit both orderings) would file the part under `x.txt`.
fn emit_after_marker(chunks: &mut [Chunk], current: usize, src: u16, marker: &str, line: u32) {
    let at = chunks[current].alloc_scratch(1);
    let from = chunks[current].alloc_scratch(1);
    let start = chunks[current].alloc_scratch(1);
    let stop = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);

    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    // Scan forward until a hit that starts a parameter.
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, from, line);
    let scan_block = chunks[current].emit_block(line);
    let (scan_loop, _) = chunks[current].emit_loop_s(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    chunks[current].emit_string_const(marker, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, from, line);
    let index_of = chunks[current].add_import("ecma:string", "indexOf");
    chunks[current].emit_call(index_of, 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, at, line);

    // No further occurrence — leave `at` negative and stop.
    chunks[current].emit_op_u16(Op::LOCAL_GET, at, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_br_if(1, line);

    // Position 0 always starts a parameter.
    chunks[current].emit_op_u16(Op::LOCAL_GET, at, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_br_if(1, line);

    // Otherwise the preceding char must be `;` (59) or a space (32).
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, at, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    let char_code_at = chunks[current].add_import("wasm:js-string", "charCodeAt");
    chunks[current].emit_call(char_code_at, 2, line);
    let prev = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, prev, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, prev, line);
    chunks[current].emit_i32_const(59, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, prev, line);
    chunks[current].emit_i32_const(32, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_br_if(1, line);
    // …or a newline (10), which is how a HEADER NAME begins: the part's own
    // `Content-Type:` is looked up with this same scan.
    chunks[current].emit_op_u16(Op::LOCAL_GET, prev, line);
    chunks[current].emit_i32_const(10, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_br_if(1, line);

    // Mid-word hit (`name=` inside `filename=`) — resume past it.
    chunks[current].emit_op_u16(Op::LOCAL_GET, at, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, from, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(scan_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(scan_block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, at, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, at, line);
    chunks[current].emit_i32_const(marker.len() as i32, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start, line);

    // Stop at the first `;` or CR after the value, whichever comes first.
    emit_next_stop(chunks, current, src, start, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stop, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stop, line);
    let substring = chunks[current].add_import("wasm:js-string", "substring");
    chunks[current].emit_call(substring, 3, line);
    let trim = chunks[current].add_import("ecma:string", "trim");
    chunks[current].emit_call(trim, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// Index of the first `;` or `\r` at or after `from`, else the string length.
fn emit_next_stop(chunks: &mut [Chunk], current: usize, src: u16, from: u16, line: u32) {
    let semi = chunks[current].alloc_scratch(1);
    let cr = chunks[current].alloc_scratch(1);
    let end = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    let len = chunks[current].add_import("wasm:js-string", "length");
    chunks[current].emit_call(len, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, end, line);

    for (marker, slot) in [(";", semi), ("\r", cr)] {
        chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
        chunks[current].emit_string_const(marker, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, from, line);
        let index_of = chunks[current].add_import("ecma:string", "indexOf");
        chunks[current].emit_call(index_of, 3, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op(Op::I32_GE_S, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, end, line);
        chunks[current].emit_op(Op::I32_LT_S, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, end, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    }

    chunks[current].emit_op_u16(Op::LOCAL_GET, end, line);
}

/// Remove the double quotes around a quoted-string parameter value.
/// Stack: [string] → [string].
///
/// RFC 7578 §5.1 tells senders to avoid quotes inside these values, so this
/// drops every `"` rather than tracking the surrounding pair.
fn emit_strip_quotes(chunks: &mut [Chunk], current: usize, line: u32) {
    let src = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, src, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    chunks[current].emit_string_const("\"", line);
    chunks[current].emit_string_const("", line);
    let replace_all = chunks[current].add_import("ecma:string", "replaceAll");
    chunks[current].emit_call(replace_all, 3, line);
}


/// This request's parsed body: `{ "fields": map, "files": map }`.
/// Stack: [] → [map].
///
/// Dispatches on `Content-Type` — urlencoded and multipart are the two forms
/// every HTML form produces — and memoises, because the body can only be read
/// once (see `http_request_env::emit_body`) and because `$_POST`/`$_FILES`,
/// WSGI's parsed form and ASP.NET's `Request.Form` must all be the same object
/// within one request.
///
/// Any other content type (JSON, XML, a raw upload) yields empty maps; those
/// bodies are read whole through `emit_body` instead, which is what PHP's
/// `php://input` and WSGI's `wsgi.input` do.
pub fn emit_parsed_body(chunks: &mut [Chunk], current: usize, line: u32) {
    super::http_request_env::emit_memoised(
        chunks,
        current,
        PARSED_BODY_CACHE_GLOBAL,
        line,
        |chunks, current, line| {
            let content_type = chunks[current].alloc_scratch(1);
            let out = chunks[current].alloc_scratch(1);

            super::http_request_env::emit_header(chunks, current, "content-type", line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, content_type, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, content_type, line);
            chunks[current].emit_op(Op::REF_IS_NULL, line);
            chunks[current].emit_if(line);
            chunks[current].emit_string_const("", line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, content_type, line);
            chunks[current].emit_end(line);

            // Match case-insensitively: `Content-Type` is a media type, and
            // its casing is not significant (RFC 9110 §8.3).
            chunks[current].emit_op_u16(Op::LOCAL_GET, content_type, line);
            let lower = chunks[current].add_import("ecma:string", "toLowerCase");
            chunks[current].emit_call(lower, 1, line);
            chunks[current].emit_string_const("multipart/form-data", line);
            let index_of = chunks[current].add_import("ecma:string", "indexOf");
            chunks[current].emit_call(index_of, 2, line);
            chunks[current].emit_i32_const(0, line);
            chunks[current].emit_op(Op::I32_GE_S, line);
            chunks[current].emit_if(line);

            super::http_request_env::emit_body(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, content_type, line);
            emit_parse_multipart(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

            chunks[current].emit_else(line);

            super::collections::emit_map_new(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
            chunks[current].emit_string_const("fields", line);

            chunks[current].emit_op_u16(Op::LOCAL_GET, content_type, line);
            let lower2 = chunks[current].add_import("ecma:string", "toLowerCase");
            chunks[current].emit_call(lower2, 1, line);
            chunks[current].emit_string_const("application/x-www-form-urlencoded", line);
            let index_of2 = chunks[current].add_import("ecma:string", "indexOf");
            chunks[current].emit_call(index_of2, 2, line);
            chunks[current].emit_i32_const(0, line);
            chunks[current].emit_op(Op::I32_GE_S, line);
            chunks[current].emit_if(line);
            super::http_request_env::emit_body(chunks, current, line);
            emit_parse_urlencoded(chunks, current, line);
            chunks[current].emit_else(line);
            super::collections::emit_map_new(chunks, current, line);
            chunks[current].emit_end(line);

            super::collections::emit_set(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);

            chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
            chunks[current].emit_string_const("files", line);
            super::collections::emit_map_new(chunks, current, line);
            super::collections::emit_set(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);

            chunks[current].emit_end(line);

            chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
        },
    );
}

const PARSED_BODY_CACHE_GLOBAL: &str = "__vybe_request_parsed_body";

/// The body's plain fields. Stack: [] → [map]. Backs PHP `$_POST`.
pub fn emit_body_fields(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_parsed_body(chunks, current, line);
    chunks[current].emit_string_const("fields", line);
    super::collections::emit_get(chunks, current, line);
}

/// The body's uploads. Stack: [] → [map]. Backs PHP `$_FILES`.
pub fn emit_body_files(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_parsed_body(chunks, current, line);
    chunks[current].emit_string_const("files", line);
    super::collections::emit_get(chunks, current, line);
}
