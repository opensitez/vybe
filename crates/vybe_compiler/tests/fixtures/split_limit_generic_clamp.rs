use vybe_compiler::primitives::strings::{SplitOptions, emit_split};
use vybe_runtime::{Chunk, Op};
fn get(c: &mut Chunk, s: u16, l: u32) {
    c.emit_op_u16(Op::LOCAL_GET, s, l);
}
fn set(c: &mut Chunk, s: u16, l: u32) {
    c.emit_op_u16(Op::LOCAL_SET, s, l);
}
pub fn emit_split_limit(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    opts: SplitOptions,
    line: u32,
) {
    if argc < 3 {
        emit_split(&mut chunks[current], line);
        return;
    }

    let base = chunks[current].alloc_scratch(6);
    let (limit, sep, s, full, keep, tail) =
        (base, base + 1, base + 2, base + 3, base + 4, base + 5);

    set(&mut chunks[current], limit, line);
    set(&mut chunks[current], sep, line);
    set(&mut chunks[current], s, line);

    get(&mut chunks[current], s, line);
    get(&mut chunks[current], sep, line);
    emit_split(&mut chunks[current], line);
    set(&mut chunks[current], full, line);

    get(&mut chunks[current], limit, line);
    chunks[current].emit_f64_const(0.0, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    if opts.negative_drops_tail {
        // `full.slice(0, max(full.length + limit, 0))`.
        get(&mut chunks[current], full, line);
        chunks[current].emit_i32_const(0, line);
        get(&mut chunks[current], full, line);
        chunks[current].emit_op(Op::ARRAY_LENGTH, line);
        get(&mut chunks[current], limit, line);
        chunks[current].emit_op(Op::F64_ADD, line);
        emit_clamp_low_zero(chunks, current, line);
        array_call(chunks, current, "slice", 3, line);
    } else {
        get(&mut chunks[current], full, line);
    }
    chunks[current].emit_else(line);
    {
        // How many pieces stay whole before the remainder is re-joined.
        get(&mut chunks[current], limit, line);
        if opts.limit_is_pieces {
            chunks[current].emit_f64_const(1.0, line);
            chunks[current].emit_op(Op::F64_SUB, line);
            emit_clamp_low_zero(chunks, current, line);
        }
        set(&mut chunks[current], keep, line);

        get(&mut chunks[current], full, line);
        get(&mut chunks[current], keep, line);
        chunks[current].emit_i32_const(0x7FFF_FFFF, line);
        array_call(chunks, current, "slice", 3, line);
        set(&mut chunks[current], tail, line);

        get(&mut chunks[current], full, line);
        chunks[current].emit_i32_const(0, line);
        get(&mut chunks[current], keep, line);
        array_call(chunks, current, "slice", 3, line);
        // `head` reuses `full`: the fully-split array is no longer needed.
        set(&mut chunks[current], full, line);

        get(&mut chunks[current], tail, line);
        chunks[current].emit_op(Op::ARRAY_LENGTH, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);
        get(&mut chunks[current], full, line);
        get(&mut chunks[current], tail, line);
        get(&mut chunks[current], sep, line);
        array_call(chunks, current, "join", 2, line);
        array_call(chunks, current, "push", 2, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);

        get(&mut chunks[current], full, line);
    }
    chunks[current].emit_end(line);
}

/// `max(n, 0)`. Stack: `[n]` → `[n]`.
fn emit_clamp_low_zero(chunks: &mut [Chunk], current: usize, line: u32) {
    let n = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], n, line);
    get(&mut chunks[current], n, line);
    chunks[current].emit_f64_const(0.0, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], n, line);
    chunks[current].emit_end(line);
}

fn array_call(chunks: &mut [Chunk], current: usize, name: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import("ecma:array", name);
    chunks[current].emit_call(idx, argc, line);
}
