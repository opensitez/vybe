use vybe_compiler::primitives::strings::{emit_index_of, emit_length, emit_substring};
use vybe_runtime::{Chunk, Op};
fn get(c: &mut Chunk, s: u16, l: u32) {
    c.emit_op_u16(Op::LOCAL_GET, s, l);
}
fn set(c: &mut Chunk, s: u16, l: u32) {
    c.emit_op_u16(Op::LOCAL_SET, s, l);
}
fn i32c(c: &mut Chunk, v: i32, l: u32) {
    c.emit_i32_const(v, l);
}
fn concat(cs: &mut [Chunk], i: usize, l: u32) {
    vybe_compiler::primitives::strings::emit_str_concat(&mut cs[i], l);
}
pub fn emit_group_digits_impl(
    chunks: &mut [Chunk],
    current: usize,
    line: u32,
    ascii_numeric: bool,
) {
    let base = chunks[current].alloc_scratch(9);
    let (dec_point, sep, s, dot, int_end, digits, frac, out, i) = (
        base,
        base + 1,
        base + 2,
        base + 3,
        base + 4,
        base + 5,
        base + 6,
        base + 7,
        base + 8,
    );
    set(&mut chunks[current], dec_point, line);
    set(&mut chunks[current], sep, line);
    set(&mut chunks[current], s, line);

    // dot = s.indexOf(".");  int_end = dot < 0 ? s.length : dot
    get(&mut chunks[current], s, line);
    chunks[current].emit_string_const(".", line);
    emit_index_of(&mut chunks[current], line);
    vybe_compiler::primitives::instructions::core_wasm::local_tee(&mut chunks[current], line, dot);
    i32c(&mut chunks[current], 0, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], s, line);
    emit_length(&mut chunks[current], line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], dot, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], int_end, line);

    // frac = dot < 0 ? "" : dec_point ++ s.slice(dot + 1)
    get(&mut chunks[current], dot, line);
    i32c(&mut chunks[current], 0, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], dec_point, line);
    get(&mut chunks[current], s, line);
    get(&mut chunks[current], dot, line);
    i32c(&mut chunks[current], 1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    i32c(&mut chunks[current], 0x7FFF_FFFF, line);
    emit_substring(&mut chunks[current], line);
    concat(chunks, current, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], frac, line);

    // A leading sign is copied out and never grouped.
    let start = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], s, line);
    i32c(&mut chunks[current], 0, line);
    i32c(&mut chunks[current], 1, line);
    emit_substring(&mut chunks[current], line);
    let first = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], first, line);
    chunks[current].emit_string_const("+-", line);
    get(&mut chunks[current], first, line);
    emit_index_of(&mut chunks[current], line);
    i32c(&mut chunks[current], 0, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    // `indexOf("")` is 0, so an empty string would read as a sign — guard on
    // the character being non-empty.
    get(&mut chunks[current], first, line);
    emit_length(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if_value(line);
    i32c(&mut chunks[current], 1, line);
    chunks[current].emit_else(line);
    i32c(&mut chunks[current], 0, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], start, line);

    get(&mut chunks[current], s, line);
    get(&mut chunks[current], start, line);
    get(&mut chunks[current], int_end, line);
    emit_substring(&mut chunks[current], line);
    set(&mut chunks[current], digits, line);

    if !ascii_numeric {
        get(&mut chunks[current], s, line);
        i32c(&mut chunks[current], 0, line);
        get(&mut chunks[current], start, line);
        emit_substring(&mut chunks[current], line);
        set(&mut chunks[current], out, line);
    }

    // for i in 0..len(digits): if i > 0 && (len - i) % 3 == 0 { out += sep }
    let n = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], digits, line);
    emit_length(&mut chunks[current], line);
    set(&mut chunks[current], n, line);
    if !ascii_numeric {
        i32c(&mut chunks[current], 0, line);
        set(&mut chunks[current], i, line);
    }

    if ascii_numeric {
        // Only for ASCII numeric formatter output; generic Unicode path remains.
        // `first` is a private primitive String used only by the completed
        // sign decision. Its last read precedes `start`/`digits`/`out`; no
        // callback can observe its slot or String identity. Reuse it for I32
        // step, retaining the exact generic 12-slot local budget.
        let step = first;
        // For positive lengths, ((n - 1) % 3) + 1 is the first group
        // size 1..3. Zero-length input produces an unused step0; its loop
        // condition is false. No branch, coercion, or additional local.
        get(&mut chunks[current], n, line);
        i32c(&mut chunks[current], 1, line);
        chunks[current].emit_op(Op::I32_SUB, line);
        i32c(&mut chunks[current], 3, line);
        chunks[current].emit_op(Op::I32_REM_S, line);
        i32c(&mut chunks[current], 1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        set(&mut chunks[current], step, line);
        // Merge the sign and first1..3 ASCII digits in one proven-safe slice.
        // n==0 has first_size0, retaining the old sign-only/empty value.
        get(&mut chunks[current], s, line);
        i32c(&mut chunks[current], 0, line);
        get(&mut chunks[current], start, line);
        get(&mut chunks[current], step, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        emit_substring(&mut chunks[current], line);
        set(&mut chunks[current], out, line);
        get(&mut chunks[current], step, line);
        set(&mut chunks[current], i, line);
        let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
        get(&mut chunks[current], i, line);
        get(&mut chunks[current], n, line);
        chunks[current].emit_op(Op::I32_GE_S, line);
        vybe_compiler::primitives::loops::emit_loop_break_if_from_i32(chunks, current, line);
        get(&mut chunks[current], out, line);
        get(&mut chunks[current], sep, line);
        concat(chunks, current, line);
        vybe_compiler::primitives::instructions::core_wasm::local_tee(
            &mut chunks[current],
            line,
            out,
        );
        get(&mut chunks[current], digits, line);
        get(&mut chunks[current], i, line);
        get(&mut chunks[current], i, line);
        i32c(&mut chunks[current], 3, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        // The substring start is already on the operand stack. Reuse its
        // computed end as the next private cursor before the pure string calls.
        chunks[current].emit_op_u16(Op::LOCAL_TEE, i, line);
        emit_substring(&mut chunks[current], line);
        concat(chunks, current, line);
        set(&mut chunks[current], out, line);
        vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    } else {
        let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
        get(&mut chunks[current], i, line);
        get(&mut chunks[current], n, line);
        chunks[current].emit_op(Op::I32_GE_S, line);
        vybe_compiler::primitives::loops::emit_loop_break_if_from_i32(chunks, current, line);

        get(&mut chunks[current], i, line);
        i32c(&mut chunks[current], 0, line);
        chunks[current].emit_op(Op::I32_GT_S, line);
        get(&mut chunks[current], n, line);
        get(&mut chunks[current], i, line);
        chunks[current].emit_op(Op::I32_SUB, line);
        i32c(&mut chunks[current], 3, line);
        chunks[current].emit_op(Op::I32_REM_S, line);
        i32c(&mut chunks[current], 0, line);
        chunks[current].emit_op(Op::I32_EQ, line);
        chunks[current].emit_op(Op::I32_AND, line);
        chunks[current].emit_if(line);
        get(&mut chunks[current], out, line);
        get(&mut chunks[current], sep, line);
        concat(chunks, current, line);
        set(&mut chunks[current], out, line);
        chunks[current].emit_end(line);

        get(&mut chunks[current], out, line);
        get(&mut chunks[current], digits, line);
        get(&mut chunks[current], i, line);
        get(&mut chunks[current], i, line);
        i32c(&mut chunks[current], 1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        // The substring start is already on the operand stack. Reuse its
        // computed end as the next private cursor before the pure string calls.
        chunks[current].emit_op_u16(Op::LOCAL_TEE, i, line);
        emit_substring(&mut chunks[current], line);
        concat(chunks, current, line);
        set(&mut chunks[current], out, line);

        vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    }

    get(&mut chunks[current], out, line);
    get(&mut chunks[current], frac, line);
    concat(chunks, current, line);
}
