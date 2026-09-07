use vybe_compiler::primitives::{collections, ops, paths, strings};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

pub fn emit_helper(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
    match name {
        "go.path_join" => emit_path_join(&mut chunks[current], argc, line),
        "go.path_clean" if argc == 1 => emit_path_clean(&mut chunks[current], line),
        "go.path_base" if argc == 1 => emit_path_base(&mut chunks[current], line),
        "go.path_ext" if argc == 1 => emit_path_ext(&mut chunks[current], line),
        "go.path_is_abs" if argc == 1 => emit_path_is_abs(&mut chunks[current], line),
        "go.path_split" if argc == 1 => emit_path_split(chunks, current, line),
        _ => return false,
    }
    true
}

fn emit_path_join(chunk: &mut Chunk, argc: u8, line: u32) {
    paths::emit_combine(chunk, argc, line);
    emit_path_clean(chunk, line);
}

fn emit_path_is_abs(chunk: &mut Chunk, line: u32) {
    let s = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, s, line);
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_string_const("/", line);
    strings::emit_index_of(chunk, line);
    chunk.emit_f64_const(0.0, line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    ops::emit_i32_to_bool(chunk, line);
}

fn emit_path_base(chunk: &mut Chunk, line: u32) {
    emit_path_clean(chunk, line);
    let clean = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, clean, line);
    chunk.emit_op_u16(Op::LOCAL_GET, clean, line);
    chunk.emit_string_const("/", line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("/", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, clean, line);
    paths::emit_file_name(chunk, line);
    chunk.emit_end(line);
}

fn emit_path_ext(chunk: &mut Chunk, line: u32) {
    emit_path_base(chunk, line);
    paths::emit_extension(chunk, line);
}

fn emit_path_split(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let p = chunk.alloc_scratch(2);
    let cut = p + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, p, line);

    chunk.emit_op_u16(Op::LOCAL_GET, p, line);
    chunk.emit_string_const("/", line);
    strings::emit_last_index_of(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, cut, line);

    chunk.emit_op_u16(Op::LOCAL_GET, cut, line);
    chunk.emit_f64_const(0.0, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    {
        chunk.emit_string_const("", line);
        chunk.emit_op_u16(Op::LOCAL_GET, p, line);
    }
    chunk.emit_else(line);
    {
        chunk.emit_op_u16(Op::LOCAL_GET, p, line);
        chunk.emit_f64_const(0.0, line);
        chunk.emit_op_u16(Op::LOCAL_GET, cut, line);
        chunk.emit_f64_const(1.0, line);
        ops::emit_dyn_add(chunk, line);
        strings::emit_substring(chunk, line);

        chunk.emit_op_u16(Op::LOCAL_GET, p, line);
        chunk.emit_op_u16(Op::LOCAL_GET, cut, line);
        chunk.emit_f64_const(1.0, line);
        ops::emit_dyn_add(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_GET, p, line);
        strings::emit_length(chunk, line);
        strings::emit_substring(chunk, line);
    }
    chunk.emit_end(line);
    collections::emit_array_pair(chunks, current, line);
}

fn emit_path_clean(chunk: &mut Chunk, line: u32) {
    let base = chunk.alloc_scratch(6);
    let input = base;
    let abs = base + 1;
    let parts = base + 2;
    let out = base + 3;
    let i = base + 4;
    let seg = base + 5;

    let split = chunk.add_import("ecma:string", "split");
    let arr_new = chunk.add_import("ecma:array", "new");
    let arr_len = chunk.add_import("ecma:array", "length");
    let arr_get = chunk.add_import("ecma:array", "get");
    let arr_push = chunk.add_import("ecma:array", "push");
    let arr_pop = chunk.add_import("ecma:array", "pop");
    let arr_join = chunk.add_import("ecma:array", "join");

    chunk.emit_op_u16(Op::LOCAL_SET, input, line);
    chunk.emit_op_u16(Op::LOCAL_GET, input, line);
    chunk.emit_string_const("", line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const(".", line);
    chunk.emit_else(line);
    {
        chunk.emit_op_u16(Op::LOCAL_GET, input, line);
        emit_path_is_abs(chunk, line);
        ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_SET, abs, line);

        chunk.emit_op_u16(Op::LOCAL_GET, input, line);
        chunk.emit_string_const("/", line);
        chunk.emit_call(split, 2, line);
        chunk.emit_op_u16(Op::LOCAL_SET, parts, line);
        chunk.emit_call(arr_new, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, out, line);
        chunk.emit_i32_const(0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, i, line);

        chunk.emit_block(line);
        chunk.emit_loop_s(line);
        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_op_u16(Op::LOCAL_GET, parts, line);
        chunk.emit_call(arr_len, 1, line);
        chunk.emit_op(Op::I32_GE_S, line);
        chunk.emit_br_if(1, line);

        chunk.emit_op_u16(Op::LOCAL_GET, parts, line);
        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_call(arr_get, 2, line);
        chunk.emit_op_u16(Op::LOCAL_SET, seg, line);

        chunk.emit_op_u16(Op::LOCAL_GET, seg, line);
        chunk.emit_string_const("..", line);
        ops::emit_dyn_eq(chunk, line);
        ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        emit_clean_dotdot(chunk, out, abs, arr_len, arr_get, arr_push, arr_pop, line);
        chunk.emit_else(line);
        emit_clean_normal_segment(chunk, out, seg, arr_push, line);
        chunk.emit_end(line);

        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_i32_const(1, line);
        chunk.emit_op(Op::I32_ADD, line);
        chunk.emit_op_u16(Op::LOCAL_SET, i, line);
        chunk.emit_br(0, line);
        chunk.emit_end(line);
        chunk.emit_end(line);

        chunk.emit_op_u16(Op::LOCAL_GET, out, line);
        chunk.emit_string_const("/", line);
        chunk.emit_call(arr_join, 2, line);
        chunk.emit_op_u16(Op::LOCAL_SET, input, line);

        chunk.emit_op_u16(Op::LOCAL_GET, abs, line);
        chunk.emit_if(line);
        chunk.emit_string_const("/", line);
        chunk.emit_op_u16(Op::LOCAL_GET, input, line);
        strings::emit_concat(chunk, 2, line);
        chunk.emit_op_u16(Op::LOCAL_SET, input, line);
        chunk.emit_end(line);

        chunk.emit_op_u16(Op::LOCAL_GET, input, line);
        chunk.emit_string_const("", line);
        ops::emit_dyn_eq(chunk, line);
        ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if_value(line);
        chunk.emit_op_u16(Op::LOCAL_GET, abs, line);
        chunk.emit_if(line);
        chunk.emit_string_const("/", line);
        chunk.emit_else(line);
        chunk.emit_string_const(".", line);
        chunk.emit_end(line);
        chunk.emit_else(line);
        chunk.emit_op_u16(Op::LOCAL_GET, input, line);
        chunk.emit_end(line);
    }
    chunk.emit_end(line);
}

fn emit_clean_dotdot(
    chunk: &mut Chunk,
    out: u16,
    abs: u16,
    arr_len: u16,
    arr_get: u16,
    arr_push: u16,
    arr_pop: u16,
    line: u32,
) {
    emit_out_can_pop(chunk, out, arr_len, arr_get, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_call(arr_pop, 1, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, abs, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_string_const("..", line);
    chunk.emit_call(arr_push, 2, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_out_can_pop(chunk: &mut Chunk, out: u16, arr_len: u16, arr_get: u16, line: u32) {
    let len = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_call(arr_len, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len, line);

    chunk.emit_op_u16(Op::LOCAL_GET, len, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_call(arr_get, 2, line);
    chunk.emit_string_const("..", line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_end(line);
}

fn emit_clean_normal_segment(chunk: &mut Chunk, out: u16, seg: u16, arr_push: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, seg, line);
    chunk.emit_string_const("", line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, seg, line);
    chunk.emit_string_const(".", line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_op_u16(Op::LOCAL_GET, seg, line);
    chunk.emit_call(arr_push, 2, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);
}
