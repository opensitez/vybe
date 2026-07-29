use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};
use vybe_compiler::compiler::instructions::{core_wasm, host};

const STATE_KEY: &str = "__state";

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        _ => panic!("push_const: no WASM-compliant encoding for {:?}", val),
    }
}

fn reserve_slot(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn emit_to_f64(chunk: &mut Chunk, line: u32) {
    host::emit(chunk, "ecma:number", "Number", 1, line);
}

fn state_key(chunk: &mut Chunk) -> u16 {
    chunk.add_constant(Value::String(STATE_KEY.into()))
}

fn emit_random_unit_from_receiver(
    chunks: &mut [Chunk],
    current: usize,
    receiver_slot: u16,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let state_slot = reserve_slot(chunk);
    let key = state_key(chunk);

    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, state_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, state_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::compiler::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op_u16(Op::LOCAL_SET, state_slot, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, state_slot, line);
    push_const(chunk, Value::F64(1103515245.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    push_const(chunk, Value::F64(12345.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    push_const(chunk, Value::F64(2147483648.0), line);
    vybe_compiler::compiler::math::emit_c_fmod(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, state_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, state_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, state_slot, line);
    push_const(chunk, Value::F64(2147483648.0), line);
    chunk.emit_op(Op::F64_DIV, line);
}

pub fn emit_random_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let seed_slot = reserve_slot(chunk);
    match argc {
        0 => {
            chunk.emit_i32_const(1, line);
            chunk.emit_op_u16(Op::LOCAL_SET, seed_slot, line);
        }
        _ => {
            for _ in 1..argc {
                chunk.emit_op(Op::DROP, line);
            }
            emit_to_f64(chunk, line);
            chunk.emit_op_u16(Op::LOCAL_SET, seed_slot, line);
        }
    }
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, seed_slot, line);
    let key = state_key(chunk);
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

pub fn emit_random_next(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let receiver_slot = reserve_slot(chunk);
    match argc {
        0 | 1 => {
            chunk.emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
            emit_random_unit_from_receiver(chunks, current, receiver_slot, line);
            push_const(&mut chunks[current], Value::F64(2147483647.0), line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op(Op::F64_FLOOR, line);
        }
        2 => {
            let max_slot = reserve_slot(chunk);
            chunk.emit_op_u16(Op::LOCAL_SET, max_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
            emit_random_unit_from_receiver(chunks, current, receiver_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, max_slot, line);
            emit_to_f64(&mut chunks[current], line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op(Op::F64_FLOOR, line);
        }
        _ => {
            let max_slot = reserve_slot(chunk);
            let min_slot = reserve_slot(chunk);
            chunk.emit_op_u16(Op::LOCAL_SET, max_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, min_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
            emit_random_unit_from_receiver(chunks, current, receiver_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, max_slot, line);
            emit_to_f64(&mut chunks[current], line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, min_slot, line);
            emit_to_f64(&mut chunks[current], line);
            chunks[current].emit_op(Op::F64_SUB, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, min_slot, line);
            emit_to_f64(&mut chunks[current], line);
            chunks[current].emit_op(Op::F64_ADD, line);
            chunks[current].emit_op(Op::F64_FLOOR, line);
        }
    }
}

pub fn emit_random_next_double(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let receiver_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    emit_random_unit_from_receiver(chunks, current, receiver_slot, line);
}

pub fn emit_random_next_bytes(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let receiver_slot = chunk.alloc_scratch(5);
    let array_slot = receiver_slot + 1;
    let len_slot = receiver_slot + 2;
    let i_slot = receiver_slot + 3;
    let value_slot = receiver_slot + 4;

    chunk.emit_op_u16(Op::LOCAL_SET, array_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, receiver_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);

    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);

    chunk.emit_block(line);
    // Structured LOOP (closed by the two `emit_end`s below), not the
    // backward-branch `emit_loop(target_offset, line)`.
    chunk.emit_loop_s(line);

    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_compiler::compiler::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::compiler::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line);

    emit_random_unit_from_receiver(chunks, current, receiver_slot, line);
    let chunk = &mut chunks[current];
    chunk.emit_f64_const(256.0, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    chunk.emit_op(Op::NULL, line);
}
