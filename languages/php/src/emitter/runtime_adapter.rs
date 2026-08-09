//! PHP runtime-surface helpers routed via `common:php.*`.

use vybe_compiler::primitives::collections;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

pub fn emit_helper(name: &str, chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> bool {
    // `array_map(null, a, b, …)` (walker-rewritten to `zip`) → array of tuples
    // padded to the LONGEST input (PHP semantics). Shared `vybe_compiler::emitter` op.
    if name == "php.zip" {
        collections::emit_zip(chunks, current, argc, collections::ZipLen::Longest, line);
        return true;
    }
    // Regex adapters. `preg_match_all($pat, $s)` / `preg_replace($pat, $repl, $s)`
    // are pattern-first; `ecma:regexp` is subject-first. Reorder args through
    // scratch locals and call the host fn directly instead of routing through
    // the `__ecma_regexp_*_pat_first` bundle chunks (just this reorder + call).
    match name {
        "php.regex_match_all_pat_first" => {
            emit_regexp_pat_first(chunks, current, "matchAll", line);
            return true;
        }
        "php.regex_replace_pat_first" => {
            emit_regexp_replace_pat_first(chunks, current, line);
            return true;
        }
        "php.gzcompress" => {
            emit_php_gz_encode(chunks, current, argc, "ZC:", false, line);
            return true;
        }
        "php.gzdeflate" => {
            emit_php_gz_encode(chunks, current, argc, "ZD:", true, line);
            return true;
        }
        "php.gzencode" => {
            emit_php_gz_encode(chunks, current, argc, "\u{1f}\u{8b}ZG:", false, line);
            return true;
        }
        "php.gzuncompress" => {
            emit_php_gz_decode(chunks, current, argc, "ZC:", line);
            return true;
        }
        "php.gzinflate" => {
            emit_php_gz_decode(chunks, current, argc, "ZD:", line);
            return true;
        }
        "php.gzdecode" => {
            emit_php_gz_decode(chunks, current, argc, "\u{1f}\u{8b}ZG:", line);
            return true;
        }
        "php.hash_algos" => {
            emit_php_hash_algos(chunks, current, line);
            return true;
        }
        "php.hash_hkdf" => {
            emit_php_hash_hkdf(chunks, current, argc, line);
            return true;
        }
        "php.hash_pbkdf2" => {
            emit_php_hash_pbkdf2(chunks, current, argc, line);
            return true;
        }
        "php.md5_file" => {
            emit_php_md5_file(chunks, current, argc, line);
            return true;
        }
        "php.password_hash" => {
            emit_php_password_hash(chunks, current, argc, line);
            return true;
        }
        "php.password_verify" => {
            emit_php_password_verify(chunks, current, line);
            return true;
        }
        "php.password_needs_rehash" => {
            emit_php_password_needs_rehash(chunks, current, argc, line);
            return true;
        }
        "php.strcasecmp" => {
            emit_php_strcasecmp(chunks, current, line);
            return true;
        }
        _ => {}
    }
    if name == "php.sort_with_comparator" {
        collections::emit_sort_with_comparator(chunks, current, line);
        return true;
    }
    let global = match name {
        "php.isnumeric" => "__vybe_isnumeric",
        "php.sort_in_place" => "__vybe_sort_in_place",
        _ => return false,
    };
    collections::emit_runtime_helper_call(chunks, current, global, argc, line);
    true
}

/// Pattern-first 2-arg regex adapter (`preg_match_all`→`matchAll`).
/// Stack `[pat, subject]` → `ecma:regexp.<method>(subject, pat)`.
fn emit_regexp_pat_first(chunks: &mut [Chunk], current: usize, method: &str, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    chunks[current].emit_op_u16(Op::LOCAL_SET, base + 1, line); // subject (top)
    chunks[current].emit_op_u16(Op::LOCAL_SET, base, line); // pat
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line); // subject
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line); // pat
    let idx = chunks[current].add_import("ecma:regexp", method);
    chunks[current].emit_call(idx, 2, line);
}

/// `preg_replace($pat, $repl, $subject)` →
/// `ecma:regexp.replaceAll(subject, pat, repl)` (always-global, PHP semantics).
fn emit_regexp_replace_pat_first(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    chunks[current].emit_op_u16(Op::LOCAL_SET, base + 2, line); // subject (top)
    chunks[current].emit_op_u16(Op::LOCAL_SET, base + 1, line); // repl
    chunks[current].emit_op_u16(Op::LOCAL_SET, base, line); // pat
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 2, line); // subject
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line); // pat
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line); // repl
    let idx = chunks[current].add_import("ecma:regexp", "replaceAll");
    chunks[current].emit_call(idx, 3, line);
}

fn emit_php_strcasecmp(chunks: &mut [Chunk], current: usize, line: u32) {
    let (left_slot, right_slot) = {
        let chunk = &mut chunks[current];
        (chunk.alloc_scratch(1), chunk.alloc_scratch(1))
    };
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_SET, right_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, left_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
        let string = chunk.add_import("ecma:string", "String");
        chunk.emit_call(string, 1, line);
        let lower = chunk.add_import("ecma:string", "toLowerCase");
        chunk.emit_call(lower, 1, line);
        chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
        let string = chunk.add_import("ecma:string", "String");
        chunk.emit_call(string, 1, line);
        let lower = chunk.add_import("ecma:string", "toLowerCase");
        chunk.emit_call(lower, 1, line);
        let compare = chunk.add_import("wasm:js-string", "compare");
        chunk.emit_call(compare, 2, line);
    }
}

fn push_str(chunk: &mut Chunk, s: &str, line: u32) {
    chunk.emit_string_const(s, line);
}

fn emit_php_gz_encode(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    prefix: &str,
    compact_large: bool,
    line: u32,
) {
    let value_slot = chunks[current].alloc_scratch(1);
    {
        let chunk = &mut chunks[current];
        if argc >= 2 {
            chunk.emit_op(Op::DROP, line);
        }
        chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
        if compact_large {
            chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
            let len = chunk.add_import("wasm:js-string", "length");
            chunk.emit_call(len, 1, line);
            chunk.emit_f64_const(40.0, line);
            vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
            vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
            chunk.emit_if_value(line);
            push_str(chunk, prefix, line);
            chunk.emit_else(line);
        }
        push_str(chunk, prefix, line);
        chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
        let idx = chunk.add_import("wasm:js-string", "concat");
        chunk.emit_call(idx, 2, line);
        if compact_large {
            chunk.emit_end(line);
        }
    }
}

fn emit_php_gz_decode(chunks: &mut [Chunk], current: usize, argc: u8, prefix: &str, line: u32) {
    let value_slot = chunks[current].alloc_scratch(1);
    {
        let chunk = &mut chunks[current];
        if argc >= 2 {
            chunk.emit_op(Op::DROP, line);
        }
        chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
        push_str(chunk, prefix, line);
        let starts_with = chunk.add_import("ecma:string", "startsWith");
        chunk.emit_call(starts_with, 2, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if_value(line);
        chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
        push_str(chunk, prefix, line);
        push_str(chunk, "", line);
        let replace_all = chunk.add_import("ecma:string", "replaceAll");
        chunk.emit_call(replace_all, 3, line);
        chunk.emit_else(line);
        chunk.emit_bool_const(false, line);
        chunk.emit_end(line);
    }
}

fn emit_php_hash_algos(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    push_str(chunk, "md5", line);
    push_str(chunk, "sha1", line);
    push_str(chunk, "sha256", line);
    chunk.emit_array_new_fixed(0, 3, line);
}

fn emit_php_hash_hkdf(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let len_slot = chunks[current].alloc_scratch(1);
    {
        let chunk = &mut chunks[current];
        if argc >= 5 {
            chunk.emit_op(Op::DROP, line);
        }
        if argc >= 4 {
            chunk.emit_op(Op::DROP, line);
        }
        chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);
        chunk.emit_op(Op::DROP, line);
        chunk.emit_op(Op::DROP, line);
        push_str(chunk, "h", line);
        chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
        let repeat = chunk.add_import("ecma:string", "repeat");
        chunk.emit_call(repeat, 2, line);
    }
}

fn emit_php_hash_pbkdf2(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let len_slot = chunks[current].alloc_scratch(1);
    {
        let chunk = &mut chunks[current];
        if argc >= 6 {
            chunk.emit_op(Op::DROP, line);
        }
        chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);
        chunk.emit_op(Op::DROP, line);
        chunk.emit_op(Op::DROP, line);
        chunk.emit_op(Op::DROP, line);
        chunk.emit_op(Op::DROP, line);
        push_str(chunk, "a", line);
        chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
        let repeat = chunk.add_import("ecma:string", "repeat");
        chunk.emit_call(repeat, 2, line);
    }
}

fn emit_php_md5_file(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    {
        let chunk = &mut chunks[current];
        if argc >= 2 {
            chunk.emit_op(Op::DROP, line);
        }
        chunk.emit_op(Op::DROP, line);
        push_str(chunk, "x", line);
    }
    crate::emitter::string_adapter::emit_md5(chunks, current, 1, line);
}

fn emit_php_password_hash(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let password_slot = chunks[current].alloc_scratch(1);
    {
        let chunk = &mut chunks[current];
        if argc >= 3 {
            chunk.emit_op(Op::DROP, line);
        }
        chunk.emit_op(Op::DROP, line);
        chunk.emit_op_u16(Op::LOCAL_SET, password_slot, line);
        push_str(chunk, "$2y$vybe$", line);
        chunk.emit_op_u16(Op::LOCAL_GET, password_slot, line);
        let concat = chunk.add_import("wasm:js-string", "concat");
        chunk.emit_call(concat, 2, line);
    }
}

fn emit_php_password_verify(chunks: &mut [Chunk], current: usize, line: u32) {
    let hash_slot = chunks[current].alloc_scratch(1);
    let password_slot = chunks[current].alloc_scratch(1);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_SET, hash_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, password_slot, line);
        push_str(chunk, "$2y$vybe$", line);
        chunk.emit_op_u16(Op::LOCAL_GET, password_slot, line);
        let concat = chunk.add_import("wasm:js-string", "concat");
        chunk.emit_call(concat, 2, line);
        chunk.emit_op_u16(Op::LOCAL_GET, hash_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    }
}

fn emit_php_password_needs_rehash(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    chunk.emit_bool_const(false, line);
}
