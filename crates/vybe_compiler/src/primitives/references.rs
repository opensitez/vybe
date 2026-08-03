use std::sync::Arc;

use crate::primitives::pointers::{CELL_KIND, REF_KIND_KEY, REF_VALUE_KEY};
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

pub fn emit_cell_new(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) {
    chunks[current].emit_struct_new(0, 0, line);

    chunks[current].emit_dup(line);
    let kind_key = chunks[current].add_constant(Value::String(Arc::from(REF_KIND_KEY)));
    chunks[current].emit_string_const(CELL_KIND, line);
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, kind_key, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_dup(line);
    let value_key = chunks[current].add_constant(Value::String(Arc::from(REF_VALUE_KEY)));
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, value_key, line);
    chunks[current].emit_op(Op::DROP, line);
}

pub fn emit_cell_new_from_local(chunks: &mut [Chunk], current: usize, local_slot: u16, line: u32) {
    emit_cell_new(chunks, current, local_slot, line);
}

pub fn emit_cell_load(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_key = chunks[current].add_constant(Value::String(Arc::from(REF_VALUE_KEY)));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, value_key, line);
}

pub fn emit_cell_store(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) {
    let value_key = chunks[current].add_constant(Value::String(Arc::from(REF_VALUE_KEY)));
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, value_key, line);
}
