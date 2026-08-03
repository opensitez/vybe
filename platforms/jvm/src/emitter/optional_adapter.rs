//! JVM `java.util.Optional` adapters.

use vybe_compiler::primitives::collections;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

pub fn emit_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    collections::emit_array_new(chunks, current, 2, line);
}

pub fn emit_of(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    collections::emit_array_new(chunks, current, 2, line);
}

pub fn emit_of_nullable(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    emit_empty(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    emit_of(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
}
