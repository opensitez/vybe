use vybe_emitter::instructions::host;

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

use super::support::stash_args;

pub fn emit_integer_of_date(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let value_slot = base;
    let date_str_slot = chunks[current].alloc_scratch(1);

    let to_string_idx = chunks[0].add_import("ecma:string", "String");
    let parse_idx = chunks[0].add_import("ecma:date", "parse");

    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, to_string_idx, line);
    chunks[current].emit(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, date_str_slot, line);

    emit_yyyymmdd_slice(chunks, current, date_str_slot, 0.0, 4.0, line);
    emit_string_const(chunks, current, "-", line);
    emit_yyyymmdd_slice(chunks, current, date_str_slot, 4.0, 6.0, line);
    emit_string_const(chunks, current, "-", line);
    emit_yyyymmdd_slice(chunks, current, date_str_slot, 6.0, 8.0, line);
    for _ in 0..4 {
        host::emit(&mut chunks[current], "wasm:js-string", "concat", 2, line);
    }

    chunks[current].emit_op_u16(Op::CALL_IMPORT, parse_idx, line);
    chunks[current].emit(1, line);
    emit_f64_const(chunks, current, 86_400_000.0, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_TRUNC, line);
    // COBOL integer dates are 1-based from 1601-01-01, but the above yields days
    // since the Unix epoch (1601-01-01 = -134774). Shift so 1601-01-01 → 1.
    emit_f64_const(chunks, current, 134_775.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
}

fn emit_yyyymmdd_slice(
    chunks: &mut [Chunk],
    current: usize,
    date_str_slot: u16,
    start: f64,
    end: f64,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, date_str_slot, line);
    emit_f64_const(chunks, current, start, line);
    emit_f64_const(chunks, current, end, line);
    host::emit(&mut chunks[current], "wasm:js-string", "substring", 3, line);
}

fn emit_string_const(chunks: &mut [Chunk], current: usize, text: &str, line: u32) {
    chunks[current].emit_string_const(text, line);
}

fn emit_f64_const(chunks: &mut [Chunk], current: usize, value: f64, line: u32) {
    chunks[current].emit_f64_const(value, line);
}
