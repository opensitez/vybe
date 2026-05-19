use std::sync::Arc;

use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

pub fn emit_cell_new(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);

    chunks[current].emit_op(Op::DUP, line);
    let kind_key = chunks[current].add_constant(Value::String(Arc::from("__ref_kind")));
    let cell_value = chunks[current].add_constant(Value::String(Arc::from("cell")));
    chunks[current].emit_op_u16(Op::CONST, cell_value, line);
    chunks[current].emit_op_u16(Op::STRUCT_SET, kind_key, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op(Op::DUP, line);
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