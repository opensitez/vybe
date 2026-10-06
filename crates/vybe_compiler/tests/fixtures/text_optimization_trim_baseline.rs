// Frozen trim emitter before private cursors and ASCII default membership.
use vybe_compiler::primitives::strings::{
    emit_length, emit_substring, emit_trim, emit_trim_end, emit_trim_start, TrimOptions,
};
use vybe_runtime::{Chunk, Op};
fn get(c: &mut Chunk, s: u16, l: u32) {
    c.emit_op_u16(Op::LOCAL_GET, s, l);
}
fn set(c: &mut Chunk, s: u16, l: u32) {
    c.emit_op_u16(Op::LOCAL_SET, s, l);
}
fn i32c(c: &mut Chunk, n: i32, l: u32) {
    c.emit_i32_const(n, l);
}
pub fn emit_trim_chars(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    opts: TrimOptions,
    line: u32,
) {
    if argc < 2 {
        match opts.default_chars {
            None => {
                // Standard behaviour — delegate.
                match (opts.left, opts.right) {
                    (true, true) => emit_trim(&mut chunks[current], line),
                    (true, false) => emit_trim_start(&mut chunks[current], line),
                    (false, true) => emit_trim_end(&mut chunks[current], line),
                    (false, false) => {}
                }
                return;
            }
            Some(defaults) => chunks[current].emit_string_const(defaults, line),
        }
    }

    let base = chunks[current].alloc_scratch(4);
    let (chars, s, start, end) = (base, base + 1, base + 2, base + 3);
    set(&mut chunks[current], chars, line);
    set(&mut chunks[current], s, line);

    i32c(&mut chunks[current], 0, line);
    set(&mut chunks[current], start, line);
    get(&mut chunks[current], s, line);
    emit_length(&mut chunks[current], line);
    set(&mut chunks[current], end, line);

    if opts.left {
        // while start < end && chars contains s[start] { start += 1 }
        let st = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
        get(&mut chunks[current], start, line);
        get(&mut chunks[current], end, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
        vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
        emit_char_in_set(chunks, current, chars, s, start, 0, line);
        chunks[current].emit_br_if(st.break_depth(0) as u32, line);
        get(&mut chunks[current], start, line);
        i32c(&mut chunks[current], 1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        set(&mut chunks[current], start, line);
        vybe_compiler::primitives::loops::emit_loop_end(chunks, current, st, line);
    }

    if opts.right {
        // while end > start && chars contains s[end-1] { end -= 1 }
        let st = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
        get(&mut chunks[current], end, line);
        get(&mut chunks[current], start, line);
        vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
        vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
        emit_char_in_set(chunks, current, chars, s, end, -1, line);
        chunks[current].emit_br_if(st.break_depth(0) as u32, line);
        get(&mut chunks[current], end, line);
        i32c(&mut chunks[current], 1, line);
        chunks[current].emit_op(Op::I32_SUB, line);
        set(&mut chunks[current], end, line);
        vybe_compiler::primitives::loops::emit_loop_end(chunks, current, st, line);
    }

    get(&mut chunks[current], s, line);
    get(&mut chunks[current], start, line);
    get(&mut chunks[current], end, line);
    emit_substring(&mut chunks[current], line);
}

/// Push i32 `1` when `chars` does NOT contain `s[idx + offset]` — the loop's
/// break condition. Uses `indexOf` on a one-character substring so a set
/// containing regex-significant characters is still matched literally.
fn emit_char_in_set(
    chunks: &mut [Chunk],
    current: usize,
    chars: u16,
    s: u16,
    idx: u16,
    offset: i32,
    line: u32,
) {
    let chunk = &mut chunks[current];
    get(chunk, chars, line);
    get(chunk, s, line);
    get(chunk, idx, line);
    if offset != 0 {
        i32c(chunk, -offset, line);
        chunk.emit_op(Op::I32_SUB, line);
    }
    let char_at = chunk.add_import("ecma:string", "charAt");
    chunk.emit_call(char_at, 2, line);
    let index_of = chunk.add_import("ecma:string", "indexOf");
    chunk.emit_call(index_of, 2, line);
    i32c(chunk, 0, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
}
