use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, Dest, ObjSource, PlainNames, ResolvedSlot, ValueSource,
};
use vybe_compiler::primitives::instructions::host;
use vybe_compiler::primitives::{
    callable, class_context, collections, convert, loops, ops, strings, tuples,
};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;
use vybe_runtime::opcode::heaptype::HT_EXTERN;

pub fn emit_helper(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
    match name {
        "go.strings.TrimPrefix" if argc == 2 => emit_trim_prefix(chunks, current, line),
        "go.strings.TrimSuffix" if argc == 2 => emit_trim_suffix(chunks, current, line),
        "go.strings.CutPrefix" if argc == 2 => emit_cut_prefix(chunks, current, line),
        "go.strings.CutSuffix" if argc == 2 => emit_cut_suffix(chunks, current, line),
        "go.strings.Cut" if argc == 2 => emit_cut(chunks, current, line),
        "go.strings.Replace" if argc == 4 => emit_replace_n(chunks, current, line),
        "go.strings.ReplaceAll" if argc == 3 => {
            host::emit(&mut chunks[current], "ecma:string", "replaceAll", 3, line)
        }
        "go.strings.ContainsRune" if argc == 2 => emit_contains_rune(chunks, current, line),
        "go.strings.ContainsAny" if argc == 2 => emit_contains_any(chunks, current, line),
        "go.strings.ContainsFunc" if argc == 2 => {
            emit_index_func(chunks, current, true, false, line)
        }
        "go.strings.IndexByte" | "go.strings.IndexRune" if argc == 2 => {
            emit_index_rune(chunks, current, false, line)
        }
        "go.strings.IndexAny" if argc == 2 => emit_index_any(chunks, current, false, line),
        "go.strings.IndexFunc" if argc == 2 => emit_index_func(chunks, current, false, false, line),
        "go.strings.LastIndexByte" if argc == 2 => emit_last_index_byte(chunks, current, line),
        "go.strings.LastIndexAny" if argc == 2 => emit_index_any(chunks, current, true, line),
        "go.strings.LastIndexFunc" if argc == 2 => {
            emit_index_func(chunks, current, false, true, line)
        }
        "go.strings.TrimLeft" if argc == 2 => emit_trim_side(chunks, current, true, line),
        "go.strings.TrimRight" if argc == 2 => emit_trim_side(chunks, current, false, line),
        "go.strings.Trim" if argc == 2 => emit_trim_cutset(chunks, current, line),
        "go.strings.EqualFold" if argc == 2 => emit_equal_fold(chunks, current, line),
        "go.strings.Count" if argc == 2 => emit_count(chunks, current, line),
        "go.strings.ToValidUTF8" if argc == 2 => emit_to_valid_utf8(chunks, current, line),
        "go.strings.Map" if argc == 2 => emit_map(chunks, current, line),
        "go.strings.Fields" if argc == 1 => emit_fields(chunks, current, None, line),
        "go.strings.FieldsFunc" if argc == 2 => emit_fields(chunks, current, Some(()), line),
        "go.strings.SplitN" if argc == 3 => emit_split_n(chunks, current, false, line),
        "go.strings.SplitAfter" if argc == 2 => emit_split_after(chunks, current, line),
        "go.strings.SplitAfterN" if argc == 3 => emit_split_n(chunks, current, true, line),
        "go.__goReplacer.Replace" | "go.strings.Replacer.Replace" if argc == 2 => {
            emit_replacer_replace(chunks, current, false, line)
        }
        "go.__goReplacer.ReplaceCascade" | "go.strings.Replacer.ReplaceCascade" if argc == 2 => {
            emit_replacer_replace(chunks, current, true, line)
        }
        "go.__goReplacer.WriteString" | "go.strings.Replacer.WriteString" if argc == 3 => {
            emit_replacer_write_string(chunks, current, line)
        }
        _ => return false,
    }
    true
}

fn emit_trim_prefix(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let s = base;
    let prefix = base + 1;
    lset(chunks, current, prefix, line);
    lset(chunks, current, s, line);
    lget(chunks, current, s, line);
    lget(chunks, current, prefix, line);
    host::emit(&mut chunks[current], "ecma:string", "startsWith", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    lget(chunks, current, s, line);
    lget(chunks, current, prefix, line);
    strings::emit_length(&mut chunks[current], line);
    lget(chunks, current, s, line);
    strings::emit_length(&mut chunks[current], line);
    host::emit(&mut chunks[current], "ecma:string", "slice", 3, line);
    chunks[current].emit_else(line);
    lget(chunks, current, s, line);
    chunks[current].emit_end(line);
}

fn emit_trim_suffix(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let s = base;
    let suffix = base + 1;
    lset(chunks, current, suffix, line);
    lset(chunks, current, s, line);
    lget(chunks, current, s, line);
    lget(chunks, current, suffix, line);
    host::emit(&mut chunks[current], "ecma:string", "endsWith", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    lget(chunks, current, s, line);
    chunks[current].emit_i32_const(0, line);
    lget(chunks, current, s, line);
    strings::emit_length(&mut chunks[current], line);
    lget(chunks, current, suffix, line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_SUB, line);
    host::emit(&mut chunks[current], "ecma:string", "slice", 3, line);
    chunks[current].emit_else(line);
    lget(chunks, current, s, line);
    chunks[current].emit_end(line);
}

fn emit_cut_prefix(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let s = base;
    let prefix = base + 1;
    lset(chunks, current, prefix, line);
    lset(chunks, current, s, line);
    lget(chunks, current, s, line);
    lget(chunks, current, prefix, line);
    host::emit(&mut chunks[current], "ecma:string", "startsWith", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    lget(chunks, current, s, line);
    lget(chunks, current, prefix, line);
    strings::emit_length(&mut chunks[current], line);
    lget(chunks, current, s, line);
    strings::emit_length(&mut chunks[current], line);
    host::emit(&mut chunks[current], "ecma:string", "slice", 3, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    lget(chunks, current, s, line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
    tuples::emit_tuple(chunks, current, 2, line);
}

fn emit_cut_suffix(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let s = base;
    let suffix = base + 1;
    lset(chunks, current, suffix, line);
    lset(chunks, current, s, line);
    lget(chunks, current, s, line);
    lget(chunks, current, suffix, line);
    host::emit(&mut chunks[current], "ecma:string", "endsWith", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    lget(chunks, current, s, line);
    chunks[current].emit_i32_const(0, line);
    lget(chunks, current, s, line);
    strings::emit_length(&mut chunks[current], line);
    lget(chunks, current, suffix, line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_SUB, line);
    host::emit(&mut chunks[current], "ecma:string", "slice", 3, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    lget(chunks, current, s, line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
    tuples::emit_tuple(chunks, current, 2, line);
}

fn emit_cut(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    let s = base;
    let sep = base + 1;
    let idx = base + 2;
    lset(chunks, current, sep, line);
    lset(chunks, current, s, line);
    lget(chunks, current, s, line);
    lget(chunks, current, sep, line);
    strings::emit_index_of(&mut chunks[current], line);
    convert::emit_to_int(&mut chunks[current], line);
    lset(chunks, current, idx, line);
    lget(chunks, current, idx, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if_value(line);
    lget(chunks, current, s, line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_else(line);
    lget(chunks, current, s, line);
    chunks[current].emit_i32_const(0, line);
    lget(chunks, current, idx, line);
    host::emit(&mut chunks[current], "ecma:string", "slice", 3, line);
    lget(chunks, current, s, line);
    lget(chunks, current, idx, line);
    lget(chunks, current, sep, line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lget(chunks, current, s, line);
    strings::emit_length(&mut chunks[current], line);
    host::emit(&mut chunks[current], "ecma:string", "slice", 3, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_end(line);
    tuples::emit_tuple(chunks, current, 3, line);
}

fn emit_replace_n(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(5);
    let s = base;
    let old = base + 1;
    let repl = base + 2;
    let n = base + 3;
    let idx = base + 4;
    lset(chunks, current, n, line);
    lset(chunks, current, repl, line);
    lset(chunks, current, old, line);
    lset(chunks, current, s, line);
    lget(chunks, current, n, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if_value(line);
    lget(chunks, current, s, line);
    lget(chunks, current, old, line);
    lget(chunks, current, repl, line);
    host::emit(&mut chunks[current], "ecma:string", "replaceAll", 3, line);
    chunks[current].emit_else(line);
    let loop_state = loops::emit_loop_start(chunks, current, line);
    lget(chunks, current, n, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);
    lget(chunks, current, s, line);
    lget(chunks, current, old, line);
    strings::emit_index_of(&mut chunks[current], line);
    convert::emit_to_int(&mut chunks[current], line);
    lset(chunks, current, idx, line);
    lget(chunks, current, idx, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(0, line);
    lset(chunks, current, n, line);
    chunks[current].emit_else(line);
    lget(chunks, current, s, line);
    chunks[current].emit_i32_const(0, line);
    lget(chunks, current, idx, line);
    host::emit(&mut chunks[current], "ecma:string", "slice", 3, line);
    lget(chunks, current, repl, line);
    strings::emit_str_concat(&mut chunks[current], line);
    lget(chunks, current, s, line);
    lget(chunks, current, idx, line);
    lget(chunks, current, old, line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lget(chunks, current, s, line);
    strings::emit_length(&mut chunks[current], line);
    host::emit(&mut chunks[current], "ecma:string", "slice", 3, line);
    strings::emit_str_concat(&mut chunks[current], line);
    lset(chunks, current, s, line);
    lget(chunks, current, n, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    lset(chunks, current, n, line);
    chunks[current].emit_end(line);
    loops::emit_loop_end(chunks, current, loop_state, line);
    lget(chunks, current, s, line);
    chunks[current].emit_end(line);
}

fn emit_contains_rune(chunks: &mut [Chunk], current: usize, line: u32) {
    let r = chunks[current].alloc_scratch(1);
    lset(chunks, current, r, line);
    lget(chunks, current, r, line);
    host::emit(
        &mut chunks[current],
        "wasm:js-string",
        "fromCodePoint",
        1,
        line,
    );
    host::emit(&mut chunks[current], "ecma:string", "includes", 2, line);
}

fn emit_index_rune(chunks: &mut [Chunk], current: usize, last: bool, line: u32) {
    let r = chunks[current].alloc_scratch(1);
    lset(chunks, current, r, line);
    lget(chunks, current, r, line);
    host::emit(
        &mut chunks[current],
        "wasm:js-string",
        "fromCodePoint",
        1,
        line,
    );
    if last {
        strings::emit_last_index_of(&mut chunks[current], line);
    } else {
        strings::emit_index_of(&mut chunks[current], line);
    }
    convert::emit_to_int(&mut chunks[current], line);
}

fn emit_last_index_byte(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_index_rune(chunks, current, true, line);
}

fn emit_contains_any(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_index_any(chunks, current, false, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

fn emit_index_any(chunks: &mut [Chunk], current: usize, last: bool, line: u32) {
    let base = chunks[current].alloc_scratch(6);
    let s = base;
    let chars = base + 1;
    let arr = base + 2;
    let i = base + 3;
    let len = base + 4;
    let result = base + 5;
    lset(chunks, current, chars, line);
    lset(chunks, current, s, line);
    lget(chunks, current, s, line);
    host::emit(&mut chunks[current], "ecma:array", "from", 1, line);
    lset(chunks, current, arr, line);
    lget(chunks, current, arr, line);
    collections::emit_len(chunks, current, line);
    convert::emit_to_int(&mut chunks[current], line);
    lset(chunks, current, len, line);
    chunks[current].emit_i32_const(-1, line);
    lset(chunks, current, result, line);
    chunks[current].emit_i32_const(0, line);
    lset(chunks, current, i, line);
    let loop_state = loops::emit_loop_start(chunks, current, line);
    lget(chunks, current, i, line);
    lget(chunks, current, len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);
    lget(chunks, current, chars, line);
    lget(chunks, current, arr, line);
    lget(chunks, current, i, line);
    collections::emit_get(chunks, current, line);
    host::emit(&mut chunks[current], "ecma:string", "includes", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    lget(chunks, current, i, line);
    lset(chunks, current, result, line);
    if !last {
        lget(chunks, current, len, line);
        lset(chunks, current, i, line);
    }
    chunks[current].emit_end(line);
    lget(chunks, current, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(chunks, current, i, line);
    loops::emit_loop_end(chunks, current, loop_state, line);
    lget(chunks, current, result, line);
}

fn emit_index_func(chunks: &mut [Chunk], current: usize, as_bool: bool, last: bool, line: u32) {
    let base = chunks[current].alloc_scratch(7);
    let s = base;
    let f = base + 1;
    let arr = base + 2;
    let i = base + 3;
    let len = base + 4;
    let result = base + 5;
    let rune = base + 6;
    lset(chunks, current, f, line);
    lset(chunks, current, s, line);
    lget(chunks, current, s, line);
    host::emit(&mut chunks[current], "ecma:array", "from", 1, line);
    lset(chunks, current, arr, line);
    lget(chunks, current, arr, line);
    collections::emit_len(chunks, current, line);
    convert::emit_to_int(&mut chunks[current], line);
    lset(chunks, current, len, line);
    chunks[current].emit_i32_const(-1, line);
    lset(chunks, current, result, line);
    chunks[current].emit_i32_const(0, line);
    lset(chunks, current, i, line);
    let loop_state = loops::emit_loop_start(chunks, current, line);
    lget(chunks, current, i, line);
    lget(chunks, current, len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);
    lget(chunks, current, arr, line);
    lget(chunks, current, i, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    host::emit(
        &mut chunks[current],
        "wasm:js-string",
        "codePointAt",
        2,
        line,
    );
    lset(chunks, current, rune, line);
    lget(chunks, current, f, line);
    let abi = class_context::module_receiver_abi(chunks);
    let recv = callable::emit_callback_receiver(&mut chunks[current], abi, line);
    lget(chunks, current, rune, line);
    callable::emit_direct_invoke_chunk(&mut chunks[current], 1 + recv, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    lget(chunks, current, i, line);
    lset(chunks, current, result, line);
    if !last {
        lget(chunks, current, len, line);
        lset(chunks, current, i, line);
    }
    chunks[current].emit_end(line);
    lget(chunks, current, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(chunks, current, i, line);
    loops::emit_loop_end(chunks, current, loop_state, line);
    lget(chunks, current, result, line);
    if as_bool {
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op(Op::I32_GE_S, line);
        ops::emit_i32_to_bool(&mut chunks[current], line);
    }
}

fn emit_trim_side(chunks: &mut [Chunk], current: usize, left: bool, line: u32) {
    let base = chunks[current].alloc_scratch(4);
    let s = base;
    let cutset = base + 1;
    let idx = base + 2;
    let len = base + 3;
    lset(chunks, current, cutset, line);
    lset(chunks, current, s, line);
    lget(chunks, current, s, line);
    strings::emit_length(&mut chunks[current], line);
    convert::emit_to_int(&mut chunks[current], line);
    lset(chunks, current, len, line);
    chunks[current].emit_i32_const(if left { 0 } else { -1 }, line);
    lset(chunks, current, idx, line);
    let loop_state = loops::emit_loop_start(chunks, current, line);
    if left {
        lget(chunks, current, idx, line);
        lget(chunks, current, len, line);
        chunks[current].emit_op(Op::I32_LT_S, line);
    } else {
        lget(chunks, current, len, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op(Op::I32_GT_S, line);
    }
    loops::emit_loop_cond_from_i32(chunks, current, line);
    lget(chunks, current, cutset, line);
    lget(chunks, current, s, line);
    if left {
        lget(chunks, current, idx, line);
        lget(chunks, current, idx, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
    } else {
        lget(chunks, current, len, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_SUB, line);
        lget(chunks, current, len, line);
    }
    host::emit(&mut chunks[current], "ecma:string", "slice", 3, line);
    host::emit(&mut chunks[current], "ecma:string", "includes", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    loops::emit_loop_cond_from_i32(chunks, current, line);
    if left {
        lget(chunks, current, idx, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        lset(chunks, current, idx, line);
    } else {
        lget(chunks, current, len, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_SUB, line);
        lset(chunks, current, len, line);
    }
    loops::emit_loop_end(chunks, current, loop_state, line);
    lget(chunks, current, s, line);
    if left {
        lget(chunks, current, idx, line);
    } else {
        chunks[current].emit_i32_const(0, line);
    }
    lget(chunks, current, len, line);
    host::emit(&mut chunks[current], "ecma:string", "slice", 3, line);
}

fn emit_trim_cutset(chunks: &mut [Chunk], current: usize, line: u32) {
    let cutset = chunks[current].alloc_scratch(1);
    lset(chunks, current, cutset, line);
    lget(chunks, current, cutset, line);
    emit_trim_side(chunks, current, true, line);
    lget(chunks, current, cutset, line);
    emit_trim_side(chunks, current, false, line);
}

fn emit_equal_fold(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:string", "toLowerCase", 1, line);
    let rhs = chunks[current].alloc_scratch(1);
    lset(chunks, current, rhs, line);
    host::emit(&mut chunks[current], "ecma:string", "toLowerCase", 1, line);
    lget(chunks, current, rhs, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

fn emit_count(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(5);
    let s = base;
    let sub = base + 1;
    let count = base + 2;
    let idx = base + 3;
    let sub_len = base + 4;
    lset(chunks, current, sub, line);
    lset(chunks, current, s, line);
    lget(chunks, current, sub, line);
    strings::emit_length(&mut chunks[current], line);
    convert::emit_to_int(&mut chunks[current], line);
    lset(chunks, current, sub_len, line);
    lget(chunks, current, sub_len, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    lget(chunks, current, s, line);
    host::emit(&mut chunks[current], "ecma:array", "from", 1, line);
    collections::emit_len(chunks, current, line);
    convert::emit_to_int(&mut chunks[current], line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(0, line);
    lset(chunks, current, count, line);
    let loop_state = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_i32_const(1, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);
    lget(chunks, current, s, line);
    lget(chunks, current, sub, line);
    strings::emit_index_of(&mut chunks[current], line);
    convert::emit_to_int(&mut chunks[current], line);
    lset(chunks, current, idx, line);
    lget(chunks, current, idx, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(0, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);
    chunks[current].emit_else(line);
    lget(chunks, current, count, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(chunks, current, count, line);
    lget(chunks, current, s, line);
    lget(chunks, current, idx, line);
    lget(chunks, current, sub_len, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lget(chunks, current, s, line);
    strings::emit_length(&mut chunks[current], line);
    host::emit(&mut chunks[current], "ecma:string", "slice", 3, line);
    lset(chunks, current, s, line);
    chunks[current].emit_end(line);
    loops::emit_loop_end(chunks, current, loop_state, line);
    lget(chunks, current, count, line);
    chunks[current].emit_end(line);
}

fn emit_to_valid_utf8(chunks: &mut [Chunk], current: usize, _line: u32) {
    let replacement = chunks[current].alloc_scratch(1);
    // Strings are already VM strings here; keep the API-shaped adapter but do
    // not duplicate UTF-8 validation machinery in Go.
    lset(chunks, current, replacement, _line);
}

fn emit_map(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(7);
    let s = base;
    let f = base + 1;
    let arr = base + 2;
    let i = base + 3;
    let len = base + 4;
    let out = base + 5;
    let mapped = base + 6;
    lset(chunks, current, s, line);
    lset(chunks, current, f, line);
    lget(chunks, current, s, line);
    host::emit(&mut chunks[current], "ecma:array", "from", 1, line);
    lset(chunks, current, arr, line);
    lget(chunks, current, arr, line);
    collections::emit_len(chunks, current, line);
    convert::emit_to_int(&mut chunks[current], line);
    lset(chunks, current, len, line);
    chunks[current].emit_string_const("", line);
    lset(chunks, current, out, line);
    chunks[current].emit_i32_const(0, line);
    lset(chunks, current, i, line);
    let loop_state = loops::emit_loop_start(chunks, current, line);
    lget(chunks, current, i, line);
    lget(chunks, current, len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);
    lget(chunks, current, f, line);
    let abi = class_context::module_receiver_abi(chunks);
    let recv = callable::emit_callback_receiver(&mut chunks[current], abi, line);
    lget(chunks, current, arr, line);
    lget(chunks, current, i, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    host::emit(
        &mut chunks[current],
        "wasm:js-string",
        "codePointAt",
        2,
        line,
    );
    callable::emit_direct_invoke_chunk(&mut chunks[current], 1 + recv, line);
    convert::emit_to_int(&mut chunks[current], line);
    lset(chunks, current, mapped, line);
    lget(chunks, current, mapped, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_if_value(line);
    lget(chunks, current, out, line);
    lget(chunks, current, mapped, line);
    host::emit(
        &mut chunks[current],
        "wasm:js-string",
        "fromCodePoint",
        1,
        line,
    );
    strings::emit_str_concat(&mut chunks[current], line);
    lset(chunks, current, out, line);
    chunks[current].emit_end(line);
    lget(chunks, current, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(chunks, current, i, line);
    loops::emit_loop_end(chunks, current, loop_state, line);
    lget(chunks, current, out, line);
}

fn emit_fields(chunks: &mut [Chunk], current: usize, with_func: Option<()>, line: u32) {
    let base = chunks[current].alloc_scratch(8);
    let s = base;
    let f = base + 1;
    let arr = base + 2;
    let i = base + 3;
    let len = base + 4;
    let out = base + 5;
    let cur = base + 6;
    let ch = base + 7;
    if with_func.is_some() {
        lset(chunks, current, f, line);
    }
    lset(chunks, current, s, line);
    lget(chunks, current, s, line);
    host::emit(&mut chunks[current], "ecma:array", "from", 1, line);
    lset(chunks, current, arr, line);
    collections::emit_array_new(chunks, current, 0, line);
    lset(chunks, current, out, line);
    chunks[current].emit_string_const("", line);
    lset(chunks, current, cur, line);
    lget(chunks, current, arr, line);
    collections::emit_len(chunks, current, line);
    convert::emit_to_int(&mut chunks[current], line);
    lset(chunks, current, len, line);
    chunks[current].emit_i32_const(0, line);
    lset(chunks, current, i, line);
    let loop_state = loops::emit_loop_start(chunks, current, line);
    lget(chunks, current, i, line);
    lget(chunks, current, len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);
    lget(chunks, current, arr, line);
    lget(chunks, current, i, line);
    collections::emit_get(chunks, current, line);
    lset(chunks, current, ch, line);
    emit_is_separator(chunks, current, ch, f, with_func.is_some(), line);
    chunks[current].emit_if_value(line);
    lget(chunks, current, cur, line);
    strings::emit_length(&mut chunks[current], line);
    convert::emit_to_int(&mut chunks[current], line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if_value(line);
    lget(chunks, current, out, line);
    lget(chunks, current, cur, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_string_const("", line);
    lset(chunks, current, cur, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    lget(chunks, current, cur, line);
    lget(chunks, current, ch, line);
    strings::emit_str_concat(&mut chunks[current], line);
    lset(chunks, current, cur, line);
    chunks[current].emit_end(line);
    lget(chunks, current, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(chunks, current, i, line);
    loops::emit_loop_end(chunks, current, loop_state, line);
    lget(chunks, current, cur, line);
    strings::emit_length(&mut chunks[current], line);
    convert::emit_to_int(&mut chunks[current], line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if_value(line);
    lget(chunks, current, out, line);
    lget(chunks, current, cur, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    lget(chunks, current, out, line);
}

fn emit_split_after(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_i32_const(-1, line);
    emit_split_n(chunks, current, true, line);
}

fn emit_split_n(chunks: &mut [Chunk], current: usize, after: bool, line: u32) {
    let base = chunks[current].alloc_scratch(7);
    let s = base;
    let sep = base + 1;
    let n = base + 2;
    let out = base + 3;
    let idx = base + 4;
    let sep_len = base + 5;
    let limit = base + 6;
    lset(chunks, current, n, line);
    lset(chunks, current, sep, line);
    lset(chunks, current, s, line);
    lget(chunks, current, n, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_else(line);
    collections::emit_array_new(chunks, current, 0, line);
    lset(chunks, current, out, line);
    lget(chunks, current, sep, line);
    strings::emit_length(&mut chunks[current], line);
    convert::emit_to_int(&mut chunks[current], line);
    lset(chunks, current, sep_len, line);
    lget(chunks, current, n, line);
    lset(chunks, current, limit, line);
    lget(chunks, current, sep_len, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    lget(chunks, current, s, line);
    host::emit(&mut chunks[current], "ecma:array", "from", 1, line);
    chunks[current].emit_else(line);
    let loop_state = loops::emit_loop_start(chunks, current, line);
    lget(chunks, current, limit, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_NE, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);
    lget(chunks, current, s, line);
    lget(chunks, current, sep, line);
    strings::emit_index_of(&mut chunks[current], line);
    convert::emit_to_int(&mut chunks[current], line);
    lset(chunks, current, idx, line);
    lget(chunks, current, idx, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(1, line);
    lset(chunks, current, limit, line);
    chunks[current].emit_else(line);
    lget(chunks, current, out, line);
    lget(chunks, current, s, line);
    chunks[current].emit_i32_const(0, line);
    lget(chunks, current, idx, line);
    if after {
        lget(chunks, current, sep_len, line);
        chunks[current].emit_op(Op::I32_ADD, line);
    }
    host::emit(&mut chunks[current], "ecma:string", "slice", 3, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(chunks, current, s, line);
    lget(chunks, current, idx, line);
    lget(chunks, current, sep_len, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lget(chunks, current, s, line);
    strings::emit_length(&mut chunks[current], line);
    host::emit(&mut chunks[current], "ecma:string", "slice", 3, line);
    lset(chunks, current, s, line);
    lget(chunks, current, limit, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if_value(line);
    lget(chunks, current, limit, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    lset(chunks, current, limit, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    loops::emit_loop_end(chunks, current, loop_state, line);
    lget(chunks, current, out, line);
    lget(chunks, current, s, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(chunks, current, out, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_replacer_replace(chunks: &mut [Chunk], current: usize, cascade: bool, line: u32) {
    let base = chunks[current].alloc_scratch(7);
    let s = base;
    let r = base + 1;
    let pairs = base + 2;
    let i = base + 3;
    let len = base + 4;
    let old = base + 5;
    let new_value = base + 6;
    lset(chunks, current, s, line);
    lset(chunks, current, r, line);
    get_field_to(&mut chunks[current], r, "pairs", pairs, line);
    lget(chunks, current, pairs, line);
    collections::emit_len(chunks, current, line);
    convert::emit_to_int(&mut chunks[current], line);
    lset(chunks, current, len, line);
    chunks[current].emit_i32_const(0, line);
    lset(chunks, current, i, line);
    let loop_state = loops::emit_loop_start(chunks, current, line);
    lget(chunks, current, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lget(chunks, current, len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);
    lget(chunks, current, pairs, line);
    lget(chunks, current, i, line);
    collections::emit_get(chunks, current, line);
    lset(chunks, current, old, line);
    lget(chunks, current, pairs, line);
    lget(chunks, current, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    collections::emit_get(chunks, current, line);
    lset(chunks, current, new_value, line);
    lget(chunks, current, old, line);
    chunks[current].emit_string_const("", line);
    ops::emit_dyn_ne(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    lget(chunks, current, s, line);
    lget(chunks, current, old, line);
    lget(chunks, current, new_value, line);
    host::emit(&mut chunks[current], "ecma:string", "replaceAll", 3, line);
    lset(chunks, current, s, line);
    if !cascade {
        chunks[current].emit_i32_const(0, line);
        lset(chunks, current, len, line);
    }
    chunks[current].emit_end(line);
    lget(chunks, current, i, line);
    chunks[current].emit_i32_const(2, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(chunks, current, i, line);
    loops::emit_loop_end(chunks, current, loop_state, line);
    lget(chunks, current, s, line);
}

fn emit_replacer_write_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    let s = base;
    let writer = base + 1;
    let replacer = base + 2;
    lset(chunks, current, s, line);
    lset(chunks, current, writer, line);
    lset(chunks, current, replacer, line);
    lget(chunks, current, replacer, line);
    lget(chunks, current, s, line);
    emit_replacer_replace(chunks, current, false, line);
    let out = chunks[current].alloc_scratch(1);
    lset(chunks, current, out, line);
    let old = chunks[current].alloc_scratch(1);
    get_field_to(&mut chunks[current], writer, "data", old, line);
    lget(chunks, current, old, line);
    lget(chunks, current, out, line);
    strings::emit_str_concat(&mut chunks[current], line);
    lset(chunks, current, old, line);
    set_local(&mut chunks[current], writer, "data", old, line);
    lget(chunks, current, out, line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
    tuples::emit_tuple(chunks, current, 2, line);
}

fn emit_is_separator(
    chunks: &mut [Chunk],
    current: usize,
    ch: u16,
    f: u16,
    with_func: bool,
    line: u32,
) {
    if with_func {
        lget(chunks, current, f, line);
        let abi = class_context::module_receiver_abi(chunks);
        let recv = callable::emit_callback_receiver(&mut chunks[current], abi, line);
        lget(chunks, current, ch, line);
        chunks[current].emit_i32_const(0, line);
        host::emit(
            &mut chunks[current],
            "wasm:js-string",
            "codePointAt",
            2,
            line,
        );
        callable::emit_direct_invoke_chunk(&mut chunks[current], 1 + recv, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
    } else {
        chunks[current].emit_string_const(" \t\n\r", line);
        lget(chunks, current, ch, line);
        host::emit(&mut chunks[current], "ecma:string", "includes", 2, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
    }
}

fn get_field_to(chunk: &mut Chunk, obj: u16, field: &str, dest: u16, line: u32) {
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(obj),
        &slot(field),
        Dest::Local(dest),
        line,
    );
}

fn set_local(chunk: &mut Chunk, obj: u16, field: &str, value: u16, line: u32) {
    class_slots::emit_class_set(
        chunk,
        ObjSource::Local(obj),
        &slot(field),
        ValueSource::Local(value),
        line,
    );
}

fn slot(field: &str) -> ResolvedSlot {
    class_slots::resolve(&ClassSlot::internal(field), &PlainNames)
}

fn lget(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
}
