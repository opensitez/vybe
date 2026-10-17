//! Bounded and unbounded zero-terminated byte scans in guest linear memory.
use vybe_runtime::{Chunk, Op};

fn get(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, slot, line);
}
fn set(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_SET, slot, line);
}
fn iconst(c: &mut Chunk, n: i32, line: u32) {
    c.emit_i32_const(n, line);
}

pub fn emit_zero_terminated_length(
    chunks: &mut [Chunk],
    current: usize,
    ptr: u16,
    limit: Option<u16>,
    line: u32,
) -> u16 {
    if limit.is_some() {
        return emit_bounded_baseline(chunks, current, ptr, limit, line);
    }
    let vector = !chunks[0].memory_is_64.first().copied().unwrap_or(false);
    let c = &mut chunks[current];
    // The original scalar emitter also reserves exactly this one scratch local.
    let len = c.alloc_scratch(1);
    iconst(c, 0, line);
    set(c, len, line);
    if !vector {
        // Memory64 callers retain the original scalar recipe exactly. The C
        // adapter also selects its existing scalar path for that memory.
        let done = c.emit_block(line);
        let (repeat, _) = c.emit_loop_s(line);
        if let Some(limit) = limit {
            get(c, len, line);
            get(c, limit, line);
            c.emit_op(Op::I32_GE_U, line);
            c.emit_br_if(1, line);
        }
        get(c, ptr, line);
        get(c, len, line);
        c.emit_op(Op::I32_ADD, line);
        c.emit_op(Op::I32_LOAD8_U, line);
        c.emit_op(Op::I32_EQZ, line);
        c.emit_br_if(1, line);
        get(c, len, line);
        iconst(c, 1, line);
        c.emit_op(Op::I32_ADD, line);
        set(c, len, line);
        c.emit_br(0, line);
        c.emit_end(line);
        c.patch_loop(repeat);
        c.emit_end(line);
        c.patch_block(done);
        return len;
    }
    let done = c.emit_block(line);
    // A short straight-line scalar prefix avoids per-iteration routing and
    // vector setup on tiny strings, with the original one scratch local.
    for offset in 0..=4 {
        if let Some(limit) = limit {
            get(c, len, line);
            get(c, limit, line);
            c.emit_op(Op::I32_GE_U, line);
            c.emit_br_if(0, line);
        }
        get(c, ptr, line);
        if offset != 0 {
            iconst(c, offset, line);
            c.emit_op(Op::I32_ADD, line);
        }
        c.emit_op(Op::I32_LOAD8_U, line);
        c.emit_op(Op::I32_EQZ, line);
        c.emit_br_if(0, line);
        iconst(c, offset + 1, line);
        set(c, len, line);
    }
    let (repeat, _) = c.emit_loop_s(line);
    if let Some(limit) = limit {
        get(c, len, line);
        get(c, limit, line);
        c.emit_op(Op::I32_GE_U, line);
        c.emit_br_if(1, line);
    }
    get(c, ptr, line);
    get(c, len, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op(Op::I32_LOAD8_U, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_br_if(1, line);
    if vector {
        // The successful scalar load proves this page exists. Standard Wasm
        // memory cannot shrink, so only the intra-page vector extent remains.
        get(c, ptr, line);
        get(c, len, line);
        c.emit_op(Op::I32_ADD, line);
        iconst(c, 65535, line);
        c.emit_op(Op::I32_AND, line);
        iconst(c, 65504, line);
        c.emit_op(Op::I32_LE_U, line);
        c.emit_if(line);
        if let Some(limit) = limit {
            get(c, limit, line);
            get(c, len, line);
            c.emit_op(Op::I32_SUB, line);
            iconst(c, 32, line);
            c.emit_op(Op::I32_GE_U, line);
            c.emit_if(line);
        }
        get(c, len, line);
        get(c, ptr, line);
        get(c, len, line);
        c.emit_op(Op::I32_ADD, line);
        c.emit_op(Op::V128_LOAD, line);
        c.emit_leb_u32(0x80, line);
        c.emit_leb_u32(0, line);
        c.emit_op(Op::V128_CONST, line);
        for _ in 0..16 {
            c.emit(0, line);
        }
        c.emit_op(Op::I8X16_EQ, line);
        c.emit_op(Op::I8X16_BITMASK, line);
        c.emit_op_u16(Op::LOCAL_TEE, len, line);
        c.emit_if_params(line, 1, 0);
        // True: the first vector found a zero. Its original length is the
        // declared block parameter; no second load or additional local.
        get(c, len, line);
        c.emit_op(Op::I32_CTZ, line);
        c.emit_op(Op::I32_ADD, line);
        set(c, len, line);
        c.emit_br(2, line);
        c.emit_else(line);
        // False: restore the original length; the low16 mask is known zero.
        set(c, len, line);
        get(c, ptr, line);
        get(c, len, line);
        c.emit_op(Op::I32_ADD, line);
        iconst(c, 16, line);
        c.emit_op(Op::I32_ADD, line);
        c.emit_op(Op::V128_LOAD, line);
        c.emit_leb_u32(0x80, line);
        c.emit_leb_u32(0, line);
        c.emit_op(Op::V128_CONST, line);
        for _ in 0..16 {
            c.emit(0, line);
        }
        c.emit_op(Op::I8X16_EQ, line);
        c.emit_op(Op::I8X16_BITMASK, line);
        // No zero in this vector advances its complete sixteen bytes.
        iconst(c, 1 << 16, line);
        c.emit_op(Op::I32_OR, line);
        c.emit_op(Op::I32_CTZ, line);
        iconst(c, 16, line);
        c.emit_op(Op::I32_ADD, line);
        get(c, len, line);
        c.emit_op(Op::I32_ADD, line);
        set(c, len, line);
        c.emit_br(2, line);
        c.emit_end(line);
        if limit.is_some() {
            c.emit_end(line);
        }
        c.emit_end(line);
    }
    get(c, len, line);
    iconst(c, 1, line);
    c.emit_op(Op::I32_ADD, line);
    set(c, len, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(repeat);
    c.emit_end(line);
    c.patch_block(done);
    len
}

fn emit_bounded_baseline(
    chunks: &mut [Chunk],
    current: usize,
    ptr: u16,
    limit: Option<u16>,
    line: u32,
) -> u16 {
    let vector = !chunks[0].memory_is_64.first().copied().unwrap_or(false);
    let c = &mut chunks[current];
    // The original scalar emitter also reserves exactly this one scratch local.
    let len = c.alloc_scratch(1);
    iconst(c, 0, line);
    set(c, len, line);
    if !vector {
        // Memory64 callers retain the original scalar recipe exactly. The C
        // adapter also selects its existing scalar path for that memory.
        let done = c.emit_block(line);
        let (repeat, _) = c.emit_loop_s(line);
        if let Some(limit) = limit {
            get(c, len, line);
            get(c, limit, line);
            c.emit_op(Op::I32_GE_U, line);
            c.emit_br_if(1, line);
        }
        get(c, ptr, line);
        get(c, len, line);
        c.emit_op(Op::I32_ADD, line);
        c.emit_op(Op::I32_LOAD8_U, line);
        c.emit_op(Op::I32_EQZ, line);
        c.emit_br_if(1, line);
        get(c, len, line);
        iconst(c, 1, line);
        c.emit_op(Op::I32_ADD, line);
        set(c, len, line);
        c.emit_br(0, line);
        c.emit_end(line);
        c.patch_loop(repeat);
        c.emit_end(line);
        c.patch_block(done);
        return len;
    }
    let done = c.emit_block(line);
    // A short straight-line scalar prefix avoids per-iteration routing and
    // vector setup on tiny strings, with the original one scratch local.
    for offset in 0..=4 {
        if let Some(limit) = limit {
            get(c, len, line);
            get(c, limit, line);
            c.emit_op(Op::I32_GE_U, line);
            c.emit_br_if(0, line);
        }
        get(c, ptr, line);
        if offset != 0 {
            iconst(c, offset, line);
            c.emit_op(Op::I32_ADD, line);
        }
        c.emit_op(Op::I32_LOAD8_U, line);
        c.emit_op(Op::I32_EQZ, line);
        c.emit_br_if(0, line);
        iconst(c, offset + 1, line);
        set(c, len, line);
    }
    let (repeat, _) = c.emit_loop_s(line);
    if let Some(limit) = limit {
        get(c, len, line);
        get(c, limit, line);
        c.emit_op(Op::I32_GE_U, line);
        c.emit_br_if(1, line);
    }
    get(c, ptr, line);
    get(c, len, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op(Op::I32_LOAD8_U, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_br_if(1, line);
    if vector {
        // The successful scalar load proves this page exists. Standard Wasm
        // memory cannot shrink, so only the intra-page vector extent remains.
        get(c, ptr, line);
        get(c, len, line);
        c.emit_op(Op::I32_ADD, line);
        iconst(c, 65535, line);
        c.emit_op(Op::I32_AND, line);
        iconst(c, 65520, line);
        c.emit_op(Op::I32_LE_U, line);
        c.emit_if(line);
        if let Some(limit) = limit {
            get(c, limit, line);
            get(c, len, line);
            c.emit_op(Op::I32_SUB, line);
            iconst(c, 16, line);
            c.emit_op(Op::I32_GE_U, line);
            c.emit_if(line);
        }
        get(c, ptr, line);
        get(c, len, line);
        c.emit_op(Op::I32_ADD, line);
        c.emit_op(Op::V128_LOAD, line);
        c.emit_leb_u32(0x80, line);
        c.emit_leb_u32(0, line);
        c.emit_op(Op::V128_CONST, line);
        for _ in 0..16 {
            c.emit(0, line);
        }
        c.emit_op(Op::I8X16_EQ, line);
        c.emit_op(Op::I8X16_BITMASK, line);
        // A sentinel gives delta16 when there is no zero. With a zero, move
        // directly to it; the next scalar probe performs the terminating read.
        iconst(c, 1 << 16, line);
        c.emit_op(Op::I32_OR, line);
        c.emit_op(Op::I32_CTZ, line);
        get(c, len, line);
        c.emit_op(Op::I32_ADD, line);
        set(c, len, line);
        c.emit_br(if limit.is_some() { 2 } else { 1 }, line);
        if limit.is_some() {
            c.emit_end(line);
        }
        c.emit_end(line);
    }
    get(c, len, line);
    iconst(c, 1, line);
    c.emit_op(Op::I32_ADD, line);
    set(c, len, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(repeat);
    c.emit_end(line);
    c.patch_block(done);
    len
}
