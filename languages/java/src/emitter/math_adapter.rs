use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;
use vybe_compiler::compiler::instructions::host;

const I32_MIN_F64: f64 = i32::MIN as f64;
const I32_MAX_F64: f64 = i32::MAX as f64;

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn emit_arithmetic_exception(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("", line);
    vybe_compiler::compiler::errors::emit_exception_new_finalize(
        &mut chunks[current],
        "ArithmeticException",
        line,
    );
    vybe_compiler::compiler::errors::emit_throw(&mut chunks[current], line);
}

fn emit_throw_if_i32_overflow(chunks: &mut [Chunk], current: usize, result: u16, line: u32) {
    get(&mut chunks[current], result, line);
    chunks[current].emit_f64_const(I32_MAX_F64, line);
    chunks[current].emit_op(Op::F64_GT, line);
    get(&mut chunks[current], result, line);
    chunks[current].emit_f64_const(I32_MIN_F64, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    emit_arithmetic_exception(chunks, current, line);
    chunks[current].emit_end(line);
}

pub fn emit_scalb(chunks: &mut [Chunk], current: usize, line: u32) {
    let scale = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], scale, line);
    set(&mut chunks[current], value, line);
    chunks[current].emit_f64_const(2.0, line);
    get(&mut chunks[current], scale, line);
    host::emit(&mut chunks[current], "ecma:math", "pow", 2, line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::F64_MUL, line);
}

pub fn emit_ulp(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:math", "abs", 1, line);
    chunks[current].emit_f64_const(f64::EPSILON, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_f64_const(f64::MIN_POSITIVE, line);
    chunks[current].emit_op(Op::F64_MAX, line);
}

pub fn emit_get_exponent(chunks: &mut [Chunk], current: usize, line: u32) {
    let abs_value = chunks[current].alloc_scratch(1);
    host::emit(&mut chunks[current], "ecma:math", "abs", 1, line);
    set(&mut chunks[current], abs_value, line);
    get(&mut chunks[current], abs_value, line);
    chunks[current].emit_f64_const(f64::MIN_POSITIVE, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(-1023.0, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], abs_value, line);
    host::emit(&mut chunks[current], "ecma:math", "log", 1, line);
    chunks[current].emit_f64_const(std::f64::consts::LN_2, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_FLOOR, line);
    chunks[current].emit_end(line);
}

pub fn emit_copy_sign(chunks: &mut [Chunk], current: usize, line: u32) {
    let sign = chunks[current].alloc_scratch(1);
    let magnitude = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], sign, line);
    set(&mut chunks[current], magnitude, line);
    get(&mut chunks[current], magnitude, line);
    host::emit(&mut chunks[current], "ecma:math", "abs", 1, line);
    chunks[current].emit_f64_const(1.0, line);
    get(&mut chunks[current], sign, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(-1.0, line);
    chunks[current].emit_else(line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op(Op::F64_MUL, line);
}

pub fn emit_next_after(chunks: &mut [Chunk], current: usize, line: u32) {
    let direction = chunks[current].alloc_scratch(1);
    let start = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], direction, line);
    set(&mut chunks[current], start, line);
    get(&mut chunks[current], direction, line);
    get(&mut chunks[current], start, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], start, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], direction, line);
    get(&mut chunks[current], start, line);
    chunks[current].emit_op(Op::F64_GT, line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], start, line);
    chunks[current].emit_f64_const(f64::EPSILON, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], start, line);
    chunks[current].emit_f64_const(f64::EPSILON, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_next_up(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_f64_const(f64::MAX, line);
    chunks[current].emit_op(Op::F64_GE, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(f64::INFINITY, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_f64_const(f64::EPSILON, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_end(line);
}

pub fn emit_next_down(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_f64_const(f64::EPSILON, line);
    chunks[current].emit_op(Op::F64_SUB, line);
}

pub fn emit_fma(chunks: &mut [Chunk], current: usize, line: u32) {
    let c = chunks[current].alloc_scratch(1);
    let b = chunks[current].alloc_scratch(1);
    let a = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], c, line);
    set(&mut chunks[current], b, line);
    set(&mut chunks[current], a, line);
    get(&mut chunks[current], a, line);
    get(&mut chunks[current], b, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    get(&mut chunks[current], c, line);
    chunks[current].emit_op(Op::F64_ADD, line);
}

pub fn emit_expm1(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:math", "exp", 1, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
}

pub fn emit_log1p(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    host::emit(&mut chunks[current], "ecma:math", "log", 1, line);
}

pub fn emit_to_degrees(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_f64_const(180.0 / std::f64::consts::PI, line);
    chunks[current].emit_op(Op::F64_MUL, line);
}

pub fn emit_to_radians(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_f64_const(std::f64::consts::PI / 180.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
}

pub fn emit_ieee_remainder(chunks: &mut [Chunk], current: usize, line: u32) {
    let divisor = chunks[current].alloc_scratch(1);
    let dividend = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], divisor, line);
    set(&mut chunks[current], dividend, line);
    get(&mut chunks[current], dividend, line);
    get(&mut chunks[current], divisor, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_NEAREST, line);
    get(&mut chunks[current], divisor, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    get(&mut chunks[current], dividend, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_op(Op::F64_NEG, line);
}

pub fn emit_add_exact(chunks: &mut [Chunk], current: usize, line: u32) {
    let rhs = chunks[current].alloc_scratch(1);
    let lhs = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], rhs, line);
    set(&mut chunks[current], lhs, line);
    get(&mut chunks[current], lhs, line);
    get(&mut chunks[current], rhs, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    set(&mut chunks[current], result, line);
    emit_throw_if_i32_overflow(chunks, current, result, line);
    get(&mut chunks[current], result, line);
}

pub fn emit_subtract_exact(chunks: &mut [Chunk], current: usize, line: u32) {
    let rhs = chunks[current].alloc_scratch(1);
    let lhs = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], rhs, line);
    set(&mut chunks[current], lhs, line);
    get(&mut chunks[current], lhs, line);
    get(&mut chunks[current], rhs, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    set(&mut chunks[current], result, line);
    emit_throw_if_i32_overflow(chunks, current, result, line);
    get(&mut chunks[current], result, line);
}

pub fn emit_multiply_exact(chunks: &mut [Chunk], current: usize, line: u32) {
    let rhs = chunks[current].alloc_scratch(1);
    let lhs = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], rhs, line);
    set(&mut chunks[current], lhs, line);
    get(&mut chunks[current], lhs, line);
    get(&mut chunks[current], rhs, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    set(&mut chunks[current], result, line);
    emit_throw_if_i32_overflow(chunks, current, result, line);
    get(&mut chunks[current], result, line);
}

pub fn emit_increment_exact(chunks: &mut [Chunk], current: usize, line: u32) {
    let result = chunks[current].alloc_scratch(1);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    set(&mut chunks[current], result, line);
    emit_throw_if_i32_overflow(chunks, current, result, line);
    get(&mut chunks[current], result, line);
}

pub fn emit_decrement_exact(chunks: &mut [Chunk], current: usize, line: u32) {
    let result = chunks[current].alloc_scratch(1);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    set(&mut chunks[current], result, line);
    emit_throw_if_i32_overflow(chunks, current, result, line);
    get(&mut chunks[current], result, line);
}

pub fn emit_negate_exact(chunks: &mut [Chunk], current: usize, line: u32) {
    let result = chunks[current].alloc_scratch(1);
    chunks[current].emit_op(Op::F64_NEG, line);
    set(&mut chunks[current], result, line);
    emit_throw_if_i32_overflow(chunks, current, result, line);
    get(&mut chunks[current], result, line);
}

pub fn emit_floor_div(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_FLOOR, line);
}

pub fn emit_floor_mod(chunks: &mut [Chunk], current: usize, line: u32) {
    let divisor = chunks[current].alloc_scratch(1);
    let dividend = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], divisor, line);
    set(&mut chunks[current], dividend, line);
    get(&mut chunks[current], dividend, line);
    get(&mut chunks[current], divisor, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_FLOOR, line);
    get(&mut chunks[current], divisor, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    get(&mut chunks[current], dividend, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_op(Op::F64_NEG, line);
}
