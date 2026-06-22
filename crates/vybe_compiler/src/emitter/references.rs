use std::sync::Arc;

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

pub fn emit_cell_new(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);

    chunks[current].emit_dup(line);
    let kind_key = chunks[current].add_constant(Value::String(Arc::from("__ref_kind")));
    chunks[current].emit_string_const("cell", line);
    chunks[current].emit_op_u16(Op::STRUCT_SET, kind_key, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_dup(line);
    let value_key = chunks[current].add_constant(Value::String(Arc::from("__value")));
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u16(Op::STRUCT_SET, value_key, line);
    chunks[current].emit_op(Op::DROP, line);
}

pub fn emit_cell_new_from_local(chunks: &mut [Chunk], current: usize, local_slot: u16, line: u32) {
    emit_cell_new(chunks, current, local_slot, line);
}

pub fn emit_cell_load(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_key = chunks[current].add_constant(Value::String(Arc::from("__value")));
    chunks[current].emit_op_u16(Op::STRUCT_GET, value_key, line);
}

pub fn emit_cell_store(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) {
    let value_key = chunks[current].add_constant(Value::String(Arc::from("__value")));
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u16(Op::STRUCT_SET, value_key, line);
}
