//! PHP `pack` / `unpack` emitter adapter.
//!
//! Format parsing and PHP result shape stay in the PHP language crate.
//! Byte/endian mechanics use `vybe_compiler::primitives::packing` so PHP shares the same
//! binary packing surface as Lua/Ruby/Python-style adapters.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}

fn emit_slot_eq_str(chunk: &mut Chunk, slot: u16, value: &str, line: u32) {
    lget(chunk, slot, line);
    let test = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(test, 1, line);
    chunk.emit_if_i32(line);
    lget(chunk, slot, line);
    chunk.emit_string_const(value, line);
    let equals = chunk.add_import("wasm:js-string", "equals");
    chunk.emit_call(equals, 2, line);
    chunk.emit_else(line);
    lget(chunk, slot, line);
    chunk.emit_string_const(value, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_end(line);
}

// PHP integer packing retains low bits; the shared checked f64 primitive
// deliberately does not own PHP casts. Keep exact integers exact until bounded.
fn emit_pack_integer(
    chunks: &mut [Chunk],
    current: usize,
    input: u16,
    width: usize,
    endian: vybe_compiler::primitives::packing::Endian,
    line: u32,
) {
    let number = alloc_local(&mut chunks[current]);
    lget(&mut chunks[current], input, line);
    call_import(chunks, current, "wasm:js-bigint", "test", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(32, line);
    lget(&mut chunks[current], input, line);
    call_import(chunks, current, "ecma:bigint", "asUintN", 2, line);
    call_import(chunks, current, "ecma:bigint", "toString", 1, line);
    call_import(chunks, current, "ecma:number", "Number", 1, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], input, line);
    call_import(chunks, current, "wasm:js-number", "test", 1, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], input, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], input, line);
    super::numeric_adapter::emit_php_floatval(chunks, current, 1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    lset(&mut chunks[current], number, line);

    // Already-bounded numeric producers go straight to the shared checked
    // primitive; only values outside its range need modulo arithmetic.
    lget(&mut chunks[current], number, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_GE, line);
    lget(&mut chunks[current], number, line);
    chunks[current].emit_f64_const(4_294_967_296.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], number, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], number, line);
    chunks[current].emit_op(Op::F64_ABS, line);
    chunks[current].emit_f64_const(f64::INFINITY, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], number, line);
    chunks[current].emit_op(Op::F64_TRUNC, line);
    lset(&mut chunks[current], number, line);
    lget(&mut chunks[current], number, line);
    lget(&mut chunks[current], number, line);
    chunks[current].emit_f64_const(4_294_967_296.0, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_FLOOR, line);
    chunks[current].emit_f64_const(4_294_967_296.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_else(line);
    // PHP's non-finite integer cast yields zero; warning policy remains separate.
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    lset(&mut chunks[current], number, line);
    match width {
        1 => vybe_compiler::primitives::packing::emit_pack_byte_from_bounded_f64_slot(
            chunks, current, number, line,
        ),
        2 => vybe_compiler::primitives::packing::emit_pack_u16_from_f64_slot(
            chunks, current, number, endian, line,
        ),
        4 => vybe_compiler::primitives::packing::emit_pack_u32_from_f64_slot(
            chunks, current, number, endian, line,
        ),
        _ => unreachable!("unsupported PHP integer pack width"),
    }
}

fn emit_concat_top(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);
}

fn emit_repeat_nul(chunk: &mut Chunk, count: usize, line: u32) {
    let s = "\0".repeat(count);
    chunk.emit_string_const(&s, line);
}

fn emit_pack_hex_string(chunks: &mut [Chunk], current: usize, hex_slot: u16, line: u32) {
    let idx_slot = alloc_local(&mut chunks[current]);
    let len_slot = alloc_local(&mut chunks[current]);
    let out_slot = alloc_local(&mut chunks[current]);
    // Collect private string pieces and join once. Rebuilding the entire
    // prefix for every byte makes large hexadecimal inputs quadratic.
    call_import(chunks, current, "ecma:array", "new", 0, line);
    lset(&mut chunks[current], out_slot, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], idx_slot, line);
    lget(&mut chunks[current], hex_slot, line);
    call_import(chunks, current, "wasm:js-string", "length", 1, line);
    lset(&mut chunks[current], len_slot, line);

    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], idx_slot, line);
    lget(&mut chunks[current], len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);

    lget(&mut chunks[current], out_slot, line);
    // The loop's indices are proven nonnegative i32 values. Use the Wasm
    // string builtin's cached UTF-16 view instead of decoding the complete
    // input again through ECMAScript slice on every byte. Preserve slice's
    // empty result when the final end index wraps at the signed-i32 limit.
    lget(&mut chunks[current], idx_slot, line);
    chunks[current].emit_i32_const(i32::MAX - 2, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], hex_slot, line);
    lget(&mut chunks[current], idx_slot, line);
    lget(&mut chunks[current], idx_slot, line);
    chunks[current].emit_i32_const(2, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    call_import(chunks, current, "wasm:js-string", "substring", 3, line);
    chunks[current].emit_end(line);
    chunks[current].emit_i32_const(16, line);
    call_import(chunks, current, "ecma:number", "parseInt", 2, line);
    call_import(chunks, current, "ecma:string", "fromCharCode", 1, line);
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    lget(&mut chunks[current], idx_slot, line);
    chunks[current].emit_i32_const(2, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], idx_slot, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);

    lget(&mut chunks[current], out_slot, line);
    chunks[current].emit_string_const("", line);
    call_import(chunks, current, "ecma:array", "join", 2, line);
}

fn emit_map_new(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "ecma:map", "new", 0, line);
}

fn emit_map_set_number_key(
    chunks: &mut [Chunk],
    current: usize,
    map_slot: u16,
    key: f64,
    value_slot: u16,
    line: u32,
) {
    lget(&mut chunks[current], map_slot, line);
    chunks[current].emit_f64_const(key, line);
    lget(&mut chunks[current], value_slot, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_map_set_string_key(
    chunks: &mut [Chunk],
    current: usize,
    map_slot: u16,
    key: &str,
    value_slot: u16,
    line: u32,
) {
    lget(&mut chunks[current], map_slot, line);
    chunks[current].emit_string_const(key, line);
    lget(&mut chunks[current], value_slot, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_unpack_one_number_key(
    chunks: &mut [Chunk],
    current: usize,
    value_slot: u16,
    key: f64,
    line: u32,
) {
    let map_slot = alloc_local(&mut chunks[current]);
    emit_map_new(chunks, current, line);
    lset(&mut chunks[current], map_slot, line);
    emit_map_set_number_key(chunks, current, map_slot, key, value_slot, line);
    lget(&mut chunks[current], map_slot, line);
}

fn emit_unpack_u32_to_slot(
    chunks: &mut [Chunk],
    current: usize,
    string_slot: u16,
    endian: vybe_compiler::primitives::packing::Endian,
    value_slot: u16,
    line: u32,
) {
    vybe_compiler::primitives::packing::emit_unpack_u32_from_string_slot_f64(
        chunks,
        current,
        string_slot,
        endian,
        line,
    );
    lset(&mut chunks[current], value_slot, line);
}

fn emit_unpack_u16_to_slot(
    chunks: &mut [Chunk],
    current: usize,
    string_slot: u16,
    endian: vybe_compiler::primitives::packing::Endian,
    value_slot: u16,
    line: u32,
) {
    vybe_compiler::primitives::packing::emit_unpack_u16_from_string_slot_f64(
        chunks,
        current,
        string_slot,
        endian,
        line,
    );
    lset(&mut chunks[current], value_slot, line);
}

pub fn emit_php_pack(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 1 {
        chunks[current].emit_string_const("", line);
        return;
    }
    emit_shared_call(chunks, current, argc, line, "pack", emit_php_pack_body);
}

fn emit_shared_call(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
    operation: &str,
    emit_body: fn(&mut Vec<Chunk>, usize, u8, u32),
) {
    let name = format!("<vybe:php:{operation}:{argc}>");
    let helper = if let Some(index) = chunks
        .iter()
        .position(|c| c.name == name && c.arity == argc)
    {
        index
    } else {
        let index = chunks.len();
        let mut body = vybe_compiler::primitives::functions::create_function_chunk(&name, argc);
        body.local_count = argc as u16;
        chunks.push(body);
        emit_body(chunks, index, argc, line);
        chunks[index].emit_op(Op::RETURN, line);
        index
    };
    if vybe_compiler::primitives::functions::try_emit_direct_helper_call(
        chunks, current, helper, argc, line,
    ) {
        return;
    }
    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        lset(&mut chunks[current], base + i as u16, line);
    }
    chunks[current].emit_op_u16(Op::REF_FUNC, helper as u16, line);
    chunks[current].emit(0, line);
    for i in 0..argc {
        lget(&mut chunks[current], base + i as u16, line);
    }
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(&mut chunks[current], argc, line);
}

fn emit_php_pack_body(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let fmt = 0;
    let first = 1;

    emit_slot_eq_str(&mut chunks[current], fmt, "C", line);
    chunks[current].emit_if_value(line);
    if argc >= 2 {
        emit_pack_integer(
            chunks,
            current,
            first,
            1,
            vybe_compiler::primitives::packing::Endian::Little,
            line,
        );
    } else {
        chunks[current].emit_string_const("", line);
    }
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "CCC", line);
    chunks[current].emit_if_value(line);
    if argc >= 4 {
        emit_pack_integer(
            chunks,
            current,
            first,
            1,
            vybe_compiler::primitives::packing::Endian::Little,
            line,
        );
        emit_pack_integer(
            chunks,
            current,
            first + 1,
            1,
            vybe_compiler::primitives::packing::Endian::Little,
            line,
        );
        emit_concat_top(chunks, current, line);
        emit_pack_integer(
            chunks,
            current,
            first + 2,
            1,
            vybe_compiler::primitives::packing::Endian::Little,
            line,
        );
        emit_concat_top(chunks, current, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "C2", line);
    chunks[current].emit_if_value(line);
    if argc >= 3 {
        emit_pack_integer(
            chunks,
            current,
            first,
            1,
            vybe_compiler::primitives::packing::Endian::Little,
            line,
        );
        emit_pack_integer(
            chunks,
            current,
            first + 1,
            1,
            vybe_compiler::primitives::packing::Endian::Little,
            line,
        );
        emit_concat_top(chunks, current, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "n", line);
    chunks[current].emit_if_value(line);
    if argc >= 2 {
        emit_pack_integer(
            chunks,
            current,
            first,
            2,
            vybe_compiler::primitives::packing::Endian::Big,
            line,
        );
    } else {
        chunks[current].emit_string_const("", line);
    }
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "v", line);
    chunks[current].emit_if_value(line);
    if argc >= 2 {
        emit_pack_integer(
            chunks,
            current,
            first,
            2,
            vybe_compiler::primitives::packing::Endian::Little,
            line,
        );
    } else {
        chunks[current].emit_string_const("", line);
    }
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "N", line);
    chunks[current].emit_if_value(line);
    if argc >= 2 {
        emit_pack_integer(
            chunks,
            current,
            first,
            4,
            vybe_compiler::primitives::packing::Endian::Big,
            line,
        );
    } else {
        chunks[current].emit_string_const("", line);
    }
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "V", line);
    chunks[current].emit_if_value(line);
    if argc >= 2 {
        emit_pack_integer(
            chunks,
            current,
            first,
            4,
            vybe_compiler::primitives::packing::Endian::Little,
            line,
        );
    } else {
        chunks[current].emit_string_const("", line);
    }
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "x", line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("\0", line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "H*", line);
    chunks[current].emit_if_value(line);
    if argc >= 2 {
        emit_pack_hex_string(chunks, current, first, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "f", line);
    chunks[current].emit_if_value(line);
    emit_repeat_nul(&mut chunks[current], 4, line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "d", line);
    chunks[current].emit_if_value(line);
    emit_repeat_nul(&mut chunks[current], 8, line);
    chunks[current].emit_else(line);

    chunks[current].emit_string_const("", line);

    for _ in 0..11 {
        chunks[current].emit_end(line);
    }
}

pub fn emit_php_unpack(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 2 {
        emit_map_new(chunks, current, line);
        return;
    }

    emit_shared_call(chunks, current, argc, line, "unpack", emit_php_unpack_body);
}

fn emit_php_unpack_body(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let fmt = 0;
    let string_slot = 1;
    let value_slot = alloc_local(&mut chunks[current]);

    emit_slot_eq_str(&mut chunks[current], fmt, "N", line);
    chunks[current].emit_if_value(line);
    emit_unpack_u32_to_slot(
        chunks,
        current,
        string_slot,
        vybe_compiler::primitives::packing::Endian::Big,
        value_slot,
        line,
    );
    emit_unpack_one_number_key(chunks, current, value_slot, 1.0, line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "V", line);
    chunks[current].emit_if_value(line);
    emit_unpack_u32_to_slot(
        chunks,
        current,
        string_slot,
        vybe_compiler::primitives::packing::Endian::Little,
        value_slot,
        line,
    );
    emit_unpack_one_number_key(chunks, current, value_slot, 1.0, line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "n", line);
    chunks[current].emit_if_value(line);
    emit_unpack_u16_to_slot(
        chunks,
        current,
        string_slot,
        vybe_compiler::primitives::packing::Endian::Big,
        value_slot,
        line,
    );
    emit_unpack_one_number_key(chunks, current, value_slot, 1.0, line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "v", line);
    chunks[current].emit_if_value(line);
    emit_unpack_u16_to_slot(
        chunks,
        current,
        string_slot,
        vybe_compiler::primitives::packing::Endian::Little,
        value_slot,
        line,
    );
    emit_unpack_one_number_key(chunks, current, value_slot, 1.0, line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "Cfirst/Csecond", line);
    chunks[current].emit_if_value(line);
    let map_slot = alloc_local(&mut chunks[current]);
    let first_slot = alloc_local(&mut chunks[current]);
    let second_slot = alloc_local(&mut chunks[current]);
    emit_map_new(chunks, current, line);
    lset(&mut chunks[current], map_slot, line);
    vybe_compiler::primitives::packing::emit_char_code_at_i32_const(
        chunks,
        current,
        string_slot,
        0,
        line,
    );
    lset(&mut chunks[current], first_slot, line);
    vybe_compiler::primitives::packing::emit_char_code_at_i32_const(
        chunks,
        current,
        string_slot,
        1,
        line,
    );
    lset(&mut chunks[current], second_slot, line);
    emit_map_set_string_key(chunks, current, map_slot, "first", first_slot, line);
    emit_map_set_string_key(chunks, current, map_slot, "second", second_slot, line);
    lget(&mut chunks[current], map_slot, line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "C*", line);
    chunks[current].emit_if_value(line);
    let star_map = alloc_local(&mut chunks[current]);
    let len_slot = alloc_local(&mut chunks[current]);
    let i_slot = alloc_local(&mut chunks[current]);
    let byte_slot = alloc_local(&mut chunks[current]);
    emit_map_new(chunks, current, line);
    lset(&mut chunks[current], star_map, line);
    lget(&mut chunks[current], string_slot, line);
    call_import(chunks, current, "wasm:js-string", "length", 1, line);
    lset(&mut chunks[current], len_slot, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], i_slot, line);
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], i_slot, line);
    lget(&mut chunks[current], len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    vybe_compiler::primitives::packing::emit_char_code_at_i32_slot(
        chunks,
        current,
        string_slot,
        i_slot,
        line,
    );
    lset(&mut chunks[current], byte_slot, line);
    lget(&mut chunks[current], star_map, line);
    lget(&mut chunks[current], i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lget(&mut chunks[current], byte_slot, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(&mut chunks[current], i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], i_slot, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    lget(&mut chunks[current], star_map, line);
    chunks[current].emit_else(line);

    emit_map_new(chunks, current, line);

    for _ in 0..6 {
        chunks[current].emit_end(line);
    }
}
