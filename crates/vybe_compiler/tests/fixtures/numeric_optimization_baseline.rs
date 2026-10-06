#![allow(dead_code)]
use vybe_compiler::primitives::Target;
use vybe_runtime::{Chunk, Op};
pub fn emit_float_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    use vybe_compiler::primitives::instructions::host;
    let v = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, v, line);

    // NaN
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    host::emit(&mut chunks[current], "ecma:number", "isNaN", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("NaN", line);
    chunks[current].emit_else(line);

    // ±Infinity
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    host::emit(&mut chunks[current], "ecma:number", "isFinite", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    vybe_compiler::primitives::math::emit_signbit(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("-Infinity", line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("Infinity", line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);

    // Negative zero: compares EQUAL to zero, so the sign bit is the only
    // witness. This is the case java lacked.
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    vybe_compiler::primitives::math::emit_signbit(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("-0.0", line);
    chunks[current].emit_else(line);
    emit_float_to_string_finite(chunks, current, v, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    emit_float_to_string_finite(chunks, current, v, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// An integral double keeps one fractional digit (`10` → `"10.0"`); anything
/// else renders as ECMA does.
fn emit_float_to_string_finite(chunks: &mut [Chunk], current: usize, v: u16, line: u32) {
    use vybe_compiler::primitives::instructions::host;
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    host::emit(&mut chunks[current], "ecma:number", "isInteger", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    chunks[current].emit_i32_const(1, line);
    host::emit(&mut chunks[current], "ecma:number", "toFixed", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    host::emit(&mut chunks[current], "ecma:number", "toString", 1, line);
    chunks[current].emit_end(line);
}
