use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    let idx = chunk.add_constant(val);
    chunk.emit_op_u16(Op::CONST, idx, line);
}

fn reserve_slot(chunk: &mut Chunk) -> u16 {
    let slot = chunk.local_count;
    chunk.local_count = slot + 1;
    slot
}

fn emit_random_unit(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = chunks[0].add_import("ecma:math", "random");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(0, line);
}

pub fn emit_random_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
}

pub fn emit_random_next(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    match argc {
        0 => {
            chunk.emit_op(Op::DROP, line);
            emit_random_unit(chunks, current, line);
            push_const(&mut chunks[current], Value::F64(2147483647.0), line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op(Op::F64_FLOOR, line);
            chunks[current].emit_op(Op::I32_FROM_F64, line);
        }
        1 => {
            let max_slot = reserve_slot(chunk);
            chunk.emit_op_u16(Op::LOCAL_SET, max_slot, line);
            chunk.emit_op(Op::DROP, line);
            chunk.emit_op(Op::DROP, line);
            emit_random_unit(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, max_slot, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op(Op::F64_FLOOR, line);
            chunks[current].emit_op(Op::I32_FROM_F64, line);
        }
        _ => {
            let max_slot = reserve_slot(chunk);
            let min_slot = reserve_slot(chunk);
            chunk.emit_op_u16(Op::LOCAL_SET, max_slot, line);
            chunk.emit_op(Op::DROP, line);
            chunk.emit_op_u16(Op::LOCAL_SET, min_slot, line);
            chunk.emit_op(Op::DROP, line);
            chunk.emit_op(Op::DROP, line);
            emit_random_unit(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, max_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, min_slot, line);
            chunks[current].emit_op(Op::F64_SUB, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, min_slot, line);
            chunks[current].emit_op(Op::F64_ADD, line);
            chunks[current].emit_op(Op::F64_FLOOR, line);
            chunks[current].emit_op(Op::I32_FROM_F64, line);
        }
    }
}

pub fn emit_random_next_double(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    emit_random_unit(chunks, current, line);
}
