//! Linear-memory byte comparison. Stack: [left address, right address, count] -> [difference].
//! Numeric storage is established by the owning adapter before selecting this emitter.
use vybe_runtime::{Chunk, Op};
fn get(c: &mut Chunk, s: u16, l: u32) {
    c.emit_op_u16(Op::LOCAL_GET, s, l);
}
fn set(c: &mut Chunk, s: u16, l: u32) {
    c.emit_op_u16(Op::LOCAL_SET, s, l);
}
fn addr(c: &mut Chunk, base: u16, index: u16, memory64: bool, l: u32) {
    get(c, base, l);
    get(c, index, l);
    c.emit_op(Op::F64_ADD, l);
    if memory64 {
        c.emit_op(Op::I64_TRUNC_F64_U, l);
    }
}
fn vector(c: &mut Chunk, base: u16, index: u16, l: u32) {
    addr(c, base, index, false, l);
    c.emit_op(Op::V128_LOAD, l);
    c.emit_leb_u32(0x80, l);
    c.emit_leb_u32(0, l);
}
fn emit_compare(
    c: &mut Chunk,
    simd: bool,
    memory64: bool,
    bounded: bool,
    nul_terminated: bool,
    l: u32,
) {
    let slots = c.alloc_scratch(if bounded { 7 } else { 6 });
    let skip = u16::from(bounded);
    let (left, right, count, index, a, b, result) = (
        slots,
        slots + 1,
        slots + 2,
        slots + 2 + skip,
        slots + 3 + skip,
        slots + 4 + skip,
        slots + 5 + skip,
    );
    if bounded {
        set(c, count, l);
    }
    set(c, right, l);
    set(c, left, l);
    c.emit_f64_const(0.0, l);
    set(c, index, l);
    c.emit_f64_const(0.0, l);
    set(c, result, l);
    let done = c.emit_block(l);
    let (repeat, _) = c.emit_loop_s(l);
    if bounded {
        get(c, index, l);
        get(c, count, l);
        c.emit_op(Op::F64_GE, l);
        c.emit_br_if(1, l);
    }
    // These scalar reads preserve early differences and establish both allocated pages.
    addr(c, left, index, memory64, l);
    c.emit_op(Op::I32_LOAD8_U, l);
    set(c, a, l);
    addr(c, right, index, memory64, l);
    c.emit_op(Op::I32_LOAD8_U, l);
    set(c, b, l);
    get(c, a, l);
    get(c, b, l);
    c.emit_op(Op::F64_NE, l);
    c.emit_if(l);
    get(c, a, l);
    get(c, b, l);
    c.emit_op(Op::F64_SUB, l);
    set(c, result, l);
    c.emit_br(2, l);
    c.emit_end(l);
    if nul_terminated {
        get(c, a, l);
        c.emit_op(Op::I32_EQZ, l);
        c.emit_br_if(1, l);
    }
    if simd {
        if bounded {
            get(c, count, l);
            get(c, index, l);
            c.emit_op(Op::F64_SUB, l);
            c.emit_i32_const(16, l);
            c.emit_op(Op::F64_GE, l);
        }
        for (position, base) in [left, right].into_iter().enumerate() {
            addr(c, base, index, false, l);
            c.emit_i32_const(65535, l);
            c.emit_op(Op::I32_AND, l);
            c.emit_i32_const(65520, l);
            c.emit_op(Op::I32_LE_U, l);
            if bounded || position != 0 {
                c.emit_op(Op::I32_AND, l);
            }
        }
        c.emit_if(l);
        vector(c, left, index, l);
        vector(c, right, index, l);
        c.emit_op(Op::I8X16_EQ, l);
        c.emit_op(Op::I8X16_BITMASK, l);
        c.emit_i32_const(65535, l);
        c.emit_op(Op::I32_XOR, l);
        if nul_terminated {
            // Stop at the first difference OR left-hand NUL. A right-hand
            // NUL without a left-hand NUL already belongs to the difference mask.
            vector(c, left, index, l);
            c.emit_op(Op::V128_CONST, l);
            for _ in 0..16 {
                c.emit(0, l);
            }
            c.emit_op(Op::I8X16_EQ, l);
            c.emit_op(Op::I8X16_BITMASK, l);
            c.emit_op(Op::I32_OR, l);
        }
        c.emit_op_u16(Op::LOCAL_TEE, a, l);
        c.emit_if(l);
        get(c, index, l);
        get(c, a, l);
        c.emit_op(Op::I32_CTZ, l);
        c.emit_op(Op::F64_ADD, l);
        set(c, index, l);
        c.emit_else(l);
        get(c, index, l);
        c.emit_i32_const(16, l);
        c.emit_op(Op::F64_ADD, l);
        set(c, index, l);
        c.emit_end(l);
        // A differing lane is re-read by the scalar path to produce its exact signed difference.
        c.emit_br(1, l);
        c.emit_end(l);
    }
    get(c, index, l);
    c.emit_i32_const(1, l);
    c.emit_op(Op::F64_ADD, l);
    set(c, index, l);
    c.emit_br(0, l);
    c.emit_end(l);
    c.patch_loop(repeat);
    c.emit_end(l);
    c.patch_block(done);
    get(c, result, l);
}

/// Both addresses have linear backing, as established by the selecting adapter.
/// Memory64 retains scalar byte reads; its address width needs a separate SIMD proof.
pub fn emit_linear_byte_compare(chunks: &mut [Chunk], current: usize, line: u32) {
    let memory64 = chunks[0].memory_is_64.first().copied().unwrap_or(false);
    emit_compare(&mut chunks[current], !memory64, memory64, true, false, line);
}

/// Compare zero-terminated linear bytes. `bounded` supplies a third count operand.
pub fn emit_linear_string_compare(chunks: &mut [Chunk], current: usize, bounded: bool, line: u32) {
    let memory64 = chunks[0].memory_is_64.first().copied().unwrap_or(false);
    emit_compare(
        &mut chunks[current],
        !memory64,
        memory64,
        bounded,
        true,
        line,
    );
}
