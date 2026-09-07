//! Python `re` surface over the ECMA `ecma:regexp` runtime — the SAME host ops
//! JS/Java/Lua drive. Python regex ≈ JS regex, so patterns pass through. No new
//! host fns. Match objects are the JS exec arrays (`m[0]` full match, `m[i]`
//! group i, `m.index` position); the walker rewrites `m.group(i)`→`m[i]` etc.
//!
//! Conventions (from java/string_adapter): `new(pattern[, flags])`→regexp;
//! `exec(regexp, str)`; `match/matchAll/search/split(str, regexp)`;
//! `replace/replaceAll(str, regexp, replacement)`.

use vybe_compiler::primitives::{ops, tuples};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use crate::emitter::adapter_util;

fn lget(c: &mut Chunk, s: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, s, line);
}
fn lset(c: &mut Chunk, s: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_SET, s, line);
}
fn call(chunks: &mut [Chunk], current: usize, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}
fn stash(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(argc as u16);
    for off in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + off, line);
    }
    base
}

/// Build a regexp from slot `pat` with optional flags string; leaves regexp on stack.
fn build_regexp(chunks: &mut [Chunk], current: usize, pat: u16, flags: Option<&str>, line: u32) {
    lget(&mut chunks[current], pat, line);
    match flags {
        Some(f) => {
            chunks[current].emit_string_const(f, line);
            call(chunks, current, "ecma:regexp", "new", 2, line);
        }
        None => call(chunks, current, "ecma:regexp", "new", 1, line),
    }
}

fn build_regexp_with_flag_slot(
    chunks: &mut [Chunk],
    current: usize,
    pat: u16,
    flags: Option<u16>,
    global: bool,
    line: u32,
) {
    lget(&mut chunks[current], pat, line);
    if let Some(flags) = flags {
        if global {
            chunks[current].emit_string_const("g", line);
            lget(&mut chunks[current], flags, line);
            call(chunks, current, "wasm:js-string", "concat", 2, line);
        } else {
            lget(&mut chunks[current], flags, line);
        }
        call(chunks, current, "ecma:regexp", "new", 2, line);
    } else if global {
        chunks[current].emit_string_const("g", line);
        call(chunks, current, "ecma:regexp", "new", 2, line);
    } else {
        call(chunks, current, "ecma:regexp", "new", 1, line);
    }
}

fn object_get_slot(chunks: &mut [Chunk], current: usize, object: u16, field: &str, line: u32) {
    lget(&mut chunks[current], object, line);
    chunks[current].emit_string_const(field, line);
    call(chunks, current, "ecma:object", "get", 2, line);
}

fn object_set_slot(
    chunks: &mut [Chunk],
    current: usize,
    object: u16,
    field: &str,
    value: u16,
    line: u32,
) {
    lget(&mut chunks[current], object, line);
    chunks[current].emit_string_const(field, line);
    lget(&mut chunks[current], value, line);
    call(chunks, current, "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// `re.search(pat, s)` → `exec(new(pat), s)` → match array or null.
pub fn emit_search(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash(chunks, current, argc, line);
    let flags = (argc >= 3).then_some(base + 2);
    build_regexp_with_flag_slot(chunks, current, base, flags, false, line);
    let r = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], r, line);
    lget(&mut chunks[current], r, line);
    lget(&mut chunks[current], base + 1, line);
    call(chunks, current, "ecma:regexp", "exec", 2, line);
}

/// `Pattern.search(s, pos, endpos)` / `Pattern.match(s, pos, endpos)`.
/// The walker sends compiled-pattern method calls here when position bounds
/// are present, so positional indices are not confused with regex flags.
pub fn emit_search_pos(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash(chunks, current, argc, line); // pat, s, pos, endpos
    let sliced = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);

    lget(&mut chunks[current], base + 1, line);
    lget(&mut chunks[current], base + 2, line);
    lget(&mut chunks[current], base + 3, line);
    call(chunks, current, "ecma:string", "slice", 3, line);
    lset(&mut chunks[current], sliced, line);

    build_regexp_with_flag_slot(chunks, current, base, None, false, line);
    let r = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], r, line);
    lget(&mut chunks[current], r, line);
    lget(&mut chunks[current], sliced, line);
    call(chunks, current, "ecma:regexp", "exec", 2, line);
    lset(&mut chunks[current], result, line);

    lget(&mut chunks[current], result, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
    chunks[current].emit_else(line);
    {
        object_get_slot(chunks, current, result, "index", line);
        lget(&mut chunks[current], base + 2, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        lset(&mut chunks[current], idx, line);
        object_set_slot(chunks, current, result, "index", idx, line);
        lget(&mut chunks[current], result, line);
    }
    chunks[current].emit_end(line);
}

/// `re.match(pat, s)` — anchored at start: `exec(new("^(?:"+pat+")"), s)`.
pub fn emit_match(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash(chunks, current, argc, line);
    let flags = (argc >= 3).then_some(base + 2);
    // pat2 = "^(?:" + pat + ")"
    chunks[current].emit_string_const("^(?:", line);
    lget(&mut chunks[current], base, line);
    call(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_string_const(")", line);
    call(chunks, current, "wasm:js-string", "concat", 2, line);
    let pat2 = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], pat2, line);
    build_regexp_with_flag_slot(chunks, current, pat2, flags, false, line);
    let r = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], r, line);
    lget(&mut chunks[current], r, line);
    lget(&mut chunks[current], base + 1, line);
    call(chunks, current, "ecma:regexp", "exec", 2, line);
}

/// `re.fullmatch(pat, s)` — anchored at both ends.
pub fn emit_fullmatch(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash(chunks, current, argc, line);
    let flags = (argc >= 3).then_some(base + 2);
    chunks[current].emit_string_const("^(?:", line);
    lget(&mut chunks[current], base, line);
    call(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_string_const(")$", line);
    call(chunks, current, "wasm:js-string", "concat", 2, line);
    let pat2 = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], pat2, line);
    build_regexp_with_flag_slot(chunks, current, pat2, flags, false, line);
    let r = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], r, line);
    lget(&mut chunks[current], r, line);
    lget(&mut chunks[current], base + 1, line);
    call(chunks, current, "ecma:regexp", "exec", 2, line);
}

/// `re.sub(pat, repl, s)` → `replaceAll(s, new(pat, "g"), repl)`.
pub fn emit_sub(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash(chunks, current, argc, line); // pat, repl, s
    lget(&mut chunks[current], base + 1, line);
    vybe_compiler::primitives::reflection::emit_is_callable(chunks, current, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    if argc >= 4 {
        lget(&mut chunks[current], base + 3, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op(Op::I32_GT_S, line);
        chunks[current].emit_if_value(line);
        emit_sub_manual(chunks, current, base, argc, true, line);
        chunks[current].emit_else(line);
        emit_sub_manual(chunks, current, base, argc, false, line);
        chunks[current].emit_end(line);
    } else {
        emit_sub_manual(chunks, current, base, argc, false, line);
    }
    chunks[current].emit_else(line);
    if argc >= 4 {
        lget(&mut chunks[current], base + 3, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op(Op::I32_GT_S, line);
        chunks[current].emit_if_value(line);
        emit_sub_manual(chunks, current, base, argc, true, line);
        chunks[current].emit_else(line);
        emit_sub_all(chunks, current, base, argc, line);
        chunks[current].emit_end(line);
    } else {
        emit_sub_all(chunks, current, base, argc, line);
    }
    chunks[current].emit_end(line);
}

fn emit_sub_all(chunks: &mut [Chunk], current: usize, base: u16, argc: u8, line: u32) {
    lget(&mut chunks[current], base + 2, line); // s
    let flags = (argc >= 5).then_some(base + 4);
    build_regexp_with_flag_slot(chunks, current, base, flags, true, line); // regexp
    lget(&mut chunks[current], base + 1, line); // repl
    call(chunks, current, "ecma:regexp", "replace", 3, line);
}

fn emit_sub_manual(
    chunks: &mut [Chunk],
    current: usize,
    base: u16,
    argc: u8,
    use_limit: bool,
    line: u32,
) {
    let flags = (argc >= 5).then_some(base + 4);
    let matches = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let replaced = chunks[current].alloc_scratch(1);
    let last = chunks[current].alloc_scratch(1);
    let m = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let piece = chunks[current].alloc_scratch(1);
    let repl_piece = chunks[current].alloc_scratch(1);

    lget(&mut chunks[current], base + 2, line);
    build_regexp_with_flag_slot(chunks, current, base, flags, true, line);
    call(chunks, current, "ecma:regexp", "matchAll", 2, line);
    lset(&mut chunks[current], matches, line);

    chunks[current].emit_string_const("", line);
    lset(&mut chunks[current], out, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], replaced, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], last, line);
    lget(&mut chunks[current], matches, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    lset(&mut chunks[current], n, line);

    let block = chunks[current].emit_block(line);
    let (lp, _) = chunks[current].emit_loop_s(line);
    lget(&mut chunks[current], i, line);
    lget(&mut chunks[current], n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    if use_limit {
        lget(&mut chunks[current], replaced, line);
        lget(&mut chunks[current], base + 3, line);
        chunks[current].emit_op(Op::I32_GE_S, line);
        chunks[current].emit_br_if(1, line);
    }

    lget(&mut chunks[current], matches, line);
    lget(&mut chunks[current], i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    lset(&mut chunks[current], m, line);
    object_get_slot(chunks, current, m, "index", line);
    lset(&mut chunks[current], idx, line);

    emit_string_slice_slot(chunks, current, base + 2, last, Some(idx), line);
    lset(&mut chunks[current], piece, line);
    lget(&mut chunks[current], out, line);
    lget(&mut chunks[current], piece, line);
    call(chunks, current, "wasm:js-string", "concat", 2, line);
    emit_replacement_piece(chunks, current, base + 1, m, repl_piece, line);
    lget(&mut chunks[current], repl_piece, line);
    call(chunks, current, "wasm:js-string", "concat", 2, line);
    lset(&mut chunks[current], out, line);

    lget(&mut chunks[current], idx, line);
    lget(&mut chunks[current], m, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    call(chunks, current, "wasm:js-string", "length", 1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], last, line);

    lget(&mut chunks[current], replaced, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], replaced, line);
    lget(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    lget(&mut chunks[current], out, line);
    emit_string_slice_slot(chunks, current, base + 2, last, None, line);
    call(chunks, current, "wasm:js-string", "concat", 2, line);
}

fn emit_replacement_piece(
    chunks: &mut [Chunk],
    current: usize,
    replacement: u16,
    match_obj: u16,
    out: u16,
    line: u32,
) {
    lget(&mut chunks[current], replacement, line);
    vybe_compiler::primitives::reflection::emit_is_callable(chunks, current, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], replacement, line);
    lget(&mut chunks[current], match_obj, line);
    vybe_compiler::primitives::callable::emit_stacked_invoke(chunks, current, 1, line);
    call(chunks, current, "ecma:string", "String", 1, line);
    lset(&mut chunks[current], out, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], replacement, line);
    lset(&mut chunks[current], out, line);
    chunks[current].emit_end(line);
}

pub fn emit_subn(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash(chunks, current, argc, line); // pat, repl, s, count?, flags?
    let flags = (argc >= 5).then_some(base + 4);
    lget(&mut chunks[current], base + 2, line);
    build_regexp_with_flag_slot(chunks, current, base, flags, true, line);
    call(chunks, current, "ecma:regexp", "matchAll", 2, line);
    let matches = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], matches, line);
    lget(&mut chunks[current], base + 2, line);
    build_regexp_with_flag_slot(chunks, current, base, flags, true, line);
    lget(&mut chunks[current], base + 1, line);
    call(chunks, current, "ecma:regexp", "replace", 3, line);
    lget(&mut chunks[current], matches, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    tuples::emit_tuple(chunks, current, 2, line);
}

/// `re.split(pat, s)` → `split(s, new(pat))`.
pub fn emit_split(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash(chunks, current, argc, line);
    if argc >= 3 {
        emit_split_maxsplit(chunks, current, base, line);
        return;
    }
    lget(&mut chunks[current], base + 1, line); // s
    build_regexp(chunks, current, base, None, line);
    call(chunks, current, "ecma:regexp", "split", 2, line);
}

fn emit_string_slice_slot(
    chunks: &mut [Chunk],
    current: usize,
    string: u16,
    start: u16,
    end: Option<u16>,
    line: u32,
) {
    lget(&mut chunks[current], string, line);
    lget(&mut chunks[current], start, line);
    match end {
        Some(end) => lget(&mut chunks[current], end, line),
        None => vybe_compiler::primitives::expressions::emit_undefined(&mut chunks[current], line),
    }
    call(chunks, current, "ecma:string", "slice", 3, line);
}

fn emit_array_push_slot(chunks: &mut [Chunk], current: usize, array: u16, value: u16, line: u32) {
    lget(&mut chunks[current], array, line);
    lget(&mut chunks[current], value, line);
    call(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_split_maxsplit(chunks: &mut [Chunk], current: usize, base: u16, line: u32) {
    lget(&mut chunks[current], base + 2, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if_value(line);
    {
        lget(&mut chunks[current], base + 1, line);
        build_regexp_with_flag_slot(chunks, current, base, None, true, line);
        call(chunks, current, "ecma:regexp", "matchAll", 2, line);
        let matches = chunks[current].alloc_scratch(1);
        lset(&mut chunks[current], matches, line);

        call(chunks, current, "ecma:array", "new", 0, line);
        let out = chunks[current].alloc_scratch(1);
        lset(&mut chunks[current], out, line);

        let i = chunks[current].alloc_scratch(1);
        let n = chunks[current].alloc_scratch(1);
        let splits = chunks[current].alloc_scratch(1);
        let last = chunks[current].alloc_scratch(1);
        let m = chunks[current].alloc_scratch(1);
        let idx = chunks[current].alloc_scratch(1);
        let piece = chunks[current].alloc_scratch(1);
        let mlen = chunks[current].alloc_scratch(1);
        let j = chunks[current].alloc_scratch(1);

        chunks[current].emit_i32_const(0, line);
        lset(&mut chunks[current], i, line);
        chunks[current].emit_i32_const(0, line);
        lset(&mut chunks[current], splits, line);
        chunks[current].emit_i32_const(0, line);
        lset(&mut chunks[current], last, line);
        lget(&mut chunks[current], matches, line);
        chunks[current].emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut chunks[current], n, line);

        let block = chunks[current].emit_block(line);
        let (lp, _) = chunks[current].emit_loop_s(line);

        lget(&mut chunks[current], i, line);
        lget(&mut chunks[current], n, line);
        chunks[current].emit_op(Op::I32_GE_S, line);
        chunks[current].emit_br_if(1, line);
        lget(&mut chunks[current], splits, line);
        lget(&mut chunks[current], base + 2, line);
        chunks[current].emit_op(Op::I32_GE_S, line);
        chunks[current].emit_br_if(1, line);

        lget(&mut chunks[current], matches, line);
        lget(&mut chunks[current], i, line);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        lset(&mut chunks[current], m, line);
        object_get_slot(chunks, current, m, "index", line);
        lset(&mut chunks[current], idx, line);

        emit_string_slice_slot(chunks, current, base + 1, last, Some(idx), line);
        lset(&mut chunks[current], piece, line);
        emit_array_push_slot(chunks, current, out, piece, line);

        lget(&mut chunks[current], m, line);
        chunks[current].emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut chunks[current], mlen, line);
        chunks[current].emit_i32_const(1, line);
        lset(&mut chunks[current], j, line);
        let captures = chunks[current].emit_block(line);
        let (cap_loop, _) = chunks[current].emit_loop_s(line);
        lget(&mut chunks[current], j, line);
        lget(&mut chunks[current], mlen, line);
        chunks[current].emit_op(Op::I32_GE_S, line);
        chunks[current].emit_br_if(1, line);
        lget(&mut chunks[current], m, line);
        lget(&mut chunks[current], j, line);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        lset(&mut chunks[current], piece, line);
        emit_array_push_slot(chunks, current, out, piece, line);
        lget(&mut chunks[current], j, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        lset(&mut chunks[current], j, line);
        chunks[current].emit_br(0, line);
        chunks[current].emit_end(line);
        chunks[current].patch_loop(cap_loop);
        chunks[current].emit_end(line);
        chunks[current].patch_block(captures);

        lget(&mut chunks[current], idx, line);
        lget(&mut chunks[current], m, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        call(chunks, current, "wasm:js-string", "length", 1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        lset(&mut chunks[current], last, line);

        lget(&mut chunks[current], splits, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        lset(&mut chunks[current], splits, line);
        lget(&mut chunks[current], i, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        lset(&mut chunks[current], i, line);
        chunks[current].emit_br(0, line);
        chunks[current].emit_end(line);
        chunks[current].patch_loop(lp);
        chunks[current].emit_end(line);
        chunks[current].patch_block(block);

        emit_string_slice_slot(chunks, current, base + 1, last, None, line);
        lset(&mut chunks[current], piece, line);
        emit_array_push_slot(chunks, current, out, piece, line);
        lget(&mut chunks[current], out, line);
    }
    chunks[current].emit_else(line);
    {
        lget(&mut chunks[current], base + 1, line);
        build_regexp(chunks, current, base, None, line);
        call(chunks, current, "ecma:regexp", "split", 2, line);
    }
    chunks[current].emit_end(line);
}

pub fn emit_finditer(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash(chunks, current, argc, line);
    lget(&mut chunks[current], base + 1, line);
    let flags = (argc >= 3).then_some(base + 2);
    build_regexp_with_flag_slot(chunks, current, base, flags, true, line);
    call(chunks, current, "ecma:regexp", "matchAll", 2, line);
}

/// `re.escape(s)` → `ecma:regexp.escape(s)`.
pub fn emit_escape(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash(chunks, current, argc, line);
    lget(&mut chunks[current], base, line);
    call(chunks, current, "ecma:regexp", "escape", 1, line);
}

/// `re.Scanner(lexicon)` returns a scanner object carrying its lexicon. The
/// current tests only construct it, but keeping the lexicon on the value gives
/// a real representation for later `.scan(...)` normalization.
pub fn emit_scanner(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash(chunks, current, argc, line);
    adapter_util::new_tagged(
        &mut chunks[current],
        "__py_re_scanner",
        &[("lexicon", base)],
        line,
    );
}

pub fn emit_match_groups(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash(chunks, current, argc, line);
    lget(&mut chunks[current], base, line);
    chunks[current].emit_i32_const(1, line);
    vybe_compiler::primitives::expressions::emit_undefined(&mut chunks[current], line);
    call(chunks, current, "ecma:array", "slice", 3, line);
    tuples::emit_tag(chunks, current, line);
}

pub fn emit_match_start(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash(chunks, current, argc, line);
    object_get_slot(chunks, current, base, "index", line);
}

pub fn emit_match_end(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash(chunks, current, argc, line);
    object_get_slot(chunks, current, base, "index", line);
    lget(&mut chunks[current], base, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    call(chunks, current, "wasm:js-string", "length", 1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
}

pub fn emit_match_lastindex(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash(chunks, current, argc, line);
    lget(&mut chunks[current], base, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    let len = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], len, line);
    lget(&mut chunks[current], len, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], len, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
}

/// `re.findall(pat, s)` → list of full matches (no group), group[1] (1 group),
/// or tuple(groups) (>1 group), via `matchAll(s, new(pat, "g"))`.
pub fn emit_findall(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash(chunks, current, argc, line); // pat, s
    lget(&mut chunks[current], base + 1, line); // s
    let flags = (argc >= 3).then_some(base + 2);
    build_regexp_with_flag_slot(chunks, current, base, flags, true, line);
    call(chunks, current, "ecma:regexp", "matchAll", 2, line);
    let matches = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], matches, line);

    let res = chunks[current].alloc_scratch(1);
    call(chunks, current, "ecma:array", "new", 0, line);
    lset(&mut chunks[current], res, line);

    let n = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let m = chunks[current].alloc_scratch(1);
    let mlen = chunks[current].alloc_scratch(1);
    lget(&mut chunks[current], matches, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    lset(&mut chunks[current], n, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], i, line);

    let block = chunks[current].emit_block(line);
    let (lp, _) = chunks[current].emit_loop_s(line);
    lget(&mut chunks[current], i, line);
    lget(&mut chunks[current], n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    // m = matches[i]; mlen = m.length
    lget(&mut chunks[current], matches, line);
    lget(&mut chunks[current], i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    lset(&mut chunks[current], m, line);
    lget(&mut chunks[current], m, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    lset(&mut chunks[current], mlen, line);

    // value = (mlen > 2) ? tuple(m[1:]) : (mlen == 2) ? m[1] : m[0]
    lget(&mut chunks[current], mlen, line);
    chunks[current].emit_i32_const(2, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if_value(line);
    {
        // tuple(m.slice(1))
        lget(&mut chunks[current], m, line);
        chunks[current].emit_i32_const(1, line);
        // §23.1.3.27 `end` = undefined means `len` — passed explicitly so this
        // import has ONE arity; a WASM import cannot have an optional argument.
        vybe_compiler::primitives::expressions::emit_undefined(&mut chunks[current], line);
        call(chunks, current, "ecma:array", "slice", 3, line);
        tuples::emit_tag(chunks, current, line);
    }
    chunks[current].emit_else(line);
    {
        lget(&mut chunks[current], mlen, line);
        chunks[current].emit_i32_const(2, line);
        chunks[current].emit_op(Op::I32_EQ, line);
        chunks[current].emit_if_value(line);
        lget(&mut chunks[current], m, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        chunks[current].emit_else(line);
        lget(&mut chunks[current], m, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_end(line);
    let val = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], val, line);

    // res.push(val)
    lget(&mut chunks[current], res, line);
    lget(&mut chunks[current], val, line);
    call(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    lget(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    lget(&mut chunks[current], res, line);
}
