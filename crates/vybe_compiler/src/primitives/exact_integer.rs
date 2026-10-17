//! Exact W64 transport between native words and the ordinary BigInt value ABI.
//!
//! Only unsigned 32-bit limbs cross Number conversions. Whole words never do.
use vybe_runtime::{Chunk, Op};

fn call(chunk: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let import = chunk.add_import(module, name);
    chunk.emit_call(import, argc, line);
}

fn align_scratch(chunk: &mut Chunk) {
    // Reclaimed value locals still have function-wide Wasm types.
    chunk.local_count = chunk.local_count.max(chunk.scratch_high_water);
}

/// Stack `[native i64]` -> `[BigInt value]`, with the declared signed reading.
/// Earlier stack operands and the caller's singleton i64 scratch are preserved.
pub fn emit_box_i64(chunks: &mut [Chunk], current: usize, signed: bool, line: u32) {
    let chunk = &mut chunks[current];
    align_scratch(chunk);
    let word = chunk.alloc_scratch(1);
    chunk.i64_locals.insert(word);
    chunk.emit_op_u16(Op::LOCAL_SET, word, line);

    chunk.emit_op_u16(Op::LOCAL_GET, word, line);
    chunk.emit_i64_const(32, line);
    chunk.emit_op(Op::I64_SHR_U, line);
    chunk.emit_op(Op::I32_WRAP_I64, line);
    chunk.emit_op(Op::F64_CONVERT_I32_U, line);
    call(chunk, "ecma:bigint", "BigInt", 1, line);
    chunk.emit_f64_const(32.0, line);
    call(chunk, "ecma:bigint", "shl", 2, line);

    chunk.emit_op_u16(Op::LOCAL_GET, word, line);
    chunk.emit_op(Op::I32_WRAP_I64, line);
    chunk.emit_op(Op::F64_CONVERT_I32_U, line);
    call(chunk, "ecma:bigint", "BigInt", 1, line);
    call(chunk, "ecma:bigint", "or", 2, line);
    if signed {
        super::bigint::emit_as_int_n(chunk, 64, line);
    }
}

/// Stack `[BigInt value]` -> `[native i64]`, retaining the low 64 bits exactly.
/// Signed and unsigned carriers have the same bit-pattern inverse. This is not
/// an implicit conversion from arbitrary floating-point or object values.
pub fn emit_unbox_i64(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    align_scratch(chunk);
    let value = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_f64_const(64.0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    call(chunk, "ecma:bigint", "asUintN", 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);

    // Normalization makes both the shifted high limb and wrapped low limb u32.
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_f64_const(32.0, line);
    call(chunk, "ecma:bigint", "shr", 2, line);
    call(chunk, "ecma:number", "Number", 1, line);
    chunk.emit_op(Op::I64_TRUNC_F64_U, line);
    chunk.emit_i64_const(32, line);
    chunk.emit_op(Op::I64_SHL, line);

    chunk.emit_f64_const(32.0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    call(chunk, "ecma:bigint", "asUintN", 2, line);
    call(chunk, "ecma:number", "Number", 1, line);
    chunk.emit_op(Op::I64_TRUNC_F64_U, line);
    chunk.emit_op(Op::I64_OR, line);
}

/// Normalize a numeric value to an exact signed/unsigned W64 carrier.
/// Existing BigInts never pass through Number; floating inputs truncate first.
pub fn emit_normalize_value(chunks: &mut [Chunk], current: usize, signed: bool, line: u32) {
    let chunk = &mut chunks[current];
    align_scratch(chunk);
    let value = chunk.alloc_scratch(1);
    crate::primitives::instructions::core_wasm::local_tee(chunk, line, value);
    call(chunk, "wasm:js-bigint", "test", 1, line);
    let exact = chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    call(chunk, "ecma:number", "Number", 1, line);
    chunk.emit_op(Op::F64_TRUNC, line);
    call(chunk, "ecma:bigint", "BigInt", 1, line);
    chunk.emit_end(line);
    chunk.patch_block(exact);
    if signed {
        super::bigint::emit_as_int_n(chunk, 64, line);
    } else {
        super::bigint::emit_as_uint_n(chunk, 64, line);
    }
}

/// Stack `[exact left, exact right]` -> exact result or native i32 predicate.
/// Only transport crosses the value ABI; the actual operation is core i64.
pub fn emit_integer_binary(
    chunks: &mut [Chunk],
    current: usize,
    operation: vybe_ast::IntOp,
    signed_result: bool,
    line: u32,
) {
    let chunk = &mut chunks[current];
    align_scratch(chunk);
    let right = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, right, line);
    emit_unbox_i64(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right, line);
    emit_unbox_i64(chunks, current, line);
    chunks[current].emit_op(
        super::bits::integer_opcode(vybe_ast::BitLane::W64, operation),
        line,
    );
    if !operation.is_comparison() {
        emit_box_i64(chunks, current, signed_result, line);
    }
}
