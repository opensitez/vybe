//! php byte-offset string operations as guest code. Strings are UTF-16 at
//! the `wasm:js-string` boundary while php offsets count UTF-8 bytes, so a
//! non-ASCII string is walked code point by code point to map byte offsets to
//! code units. A slice through the middle of a code point yields U+FFFD for
//! each undecodable piece: one per stray continuation byte at the start, one
//! for a truncated sequence at the end.
use vybe_runtime::{Chunk, opcode::Op};

const SUBSTR_HELPER: &str = "__php_substr_bytes";
const REPLACEMENT: &str = "\u{FFFD}";

fn import(chunk: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let index = chunk.add_import(module, name);
    chunk.emit_call(index, argc, line);
}

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

/// Stack: `[value] -> [f64]` — php's integer conversion of an offset, as an
/// integral f64 (NaN and non-numeric strings are 0).
fn emit_offset_integer(chunk: &mut Chunk, line: u32) {
    import(chunk, "ecma:number", "Number", 1, line);
    import(chunk, "wasm:js-number", "toF64", 1, line);
    chunk.emit_op(Op::I64_TRUNC_SAT_F64_S, line);
    chunk.emit_op(Op::F64_CONVERT_I64_S, line);
}

/// `slot = clamp(slot, low, high)` on f64 locals.
fn emit_clamp(chunk: &mut Chunk, slot: u16, low: u16, high: u16, line: u32) {
    get(chunk, slot, line);
    get(chunk, low, line);
    chunk.emit_op(Op::F64_MAX, line);
    get(chunk, high, line);
    chunk.emit_op(Op::F64_MIN, line);
    set(chunk, slot, line);
}

/// `count` U+FFFD characters, concatenated onto the string in `out`.
fn emit_append_replacements(chunk: &mut Chunk, out: u16, count: u16, line: u32) {
    let block = chunk.emit_block(line);
    let (lp, _) = chunk.emit_loop_s(line);
    get(chunk, count, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_LE, line);
    chunk.emit_br_if(1, line);
    get(chunk, out, line);
    chunk.emit_string_const(REPLACEMENT, line);
    import(chunk, "wasm:js-string", "concat", 2, line);
    set(chunk, out, line);
    get(chunk, count, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_SUB, line);
    set(chunk, count, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(lp);
    chunk.emit_end(line);
    chunk.patch_block(block);
}

/// `__php_substr_bytes(string, offset, length|null)`.
fn ensure_substr_helper(chunks: &mut Vec<Chunk>) -> usize {
    if let Some(index) = chunks.iter().position(|chunk| chunk.name == SUBSTR_HELPER) {
        return index;
    }
    let index = chunks.len();
    let mut helper = Chunk::new(SUBSTR_HELPER);
    helper.arity = 3;
    helper.local_count = 3;
    chunks.push(helper);
    let line = 0;
    let c = &mut chunks[index];
    let (text, offset_value, length_value) = (0u16, 1u16, 2u16);
    let size = c.alloc_scratch(1);
    let zero = c.alloc_scratch(1);
    let start = c.alloc_scratch(1);
    let end = c.alloc_scratch(1);
    let units = c.alloc_scratch(1);
    let unit = c.alloc_scratch(1);
    let byte = c.alloc_scratch(1);
    let width = c.alloc_scratch(1);
    let span = c.alloc_scratch(1);
    let begin = c.alloc_scratch(1);
    let finish = c.alloc_scratch(1);
    let lead = c.alloc_scratch(1);
    let tail = c.alloc_scratch(1);
    let out = c.alloc_scratch(1);
    let cp_end = c.alloc_scratch(1);

    get(c, text, line);
    import(c, "wasm:text-encoder", "measureStringAsUTF8", 1, line);
    c.emit_op(Op::F64_FROM_I32, line);
    set(c, size, line);
    c.emit_f64_const(0.0, line);
    set(c, zero, line);

    // start: a negative offset counts from the end.
    get(c, offset_value, line);
    emit_offset_integer(c, line);
    set(c, start, line);
    get(c, start, line);
    c.emit_f64_const(0.0, line);
    c.emit_op(Op::F64_LT, line);
    c.emit_if(line);
    get(c, size, line);
    get(c, start, line);
    c.emit_op(Op::F64_ADD, line);
    set(c, start, line);
    c.emit_end(line);
    emit_clamp(c, start, zero, size, line);

    // end: absent length means the rest; a negative length stops that many
    // bytes before the end.
    get(c, length_value, line);
    c.emit_op(Op::REF_IS_NULL, line);
    get(c, length_value, line);
    import(c, "wasm:js-undefined", "test", 1, line);
    c.emit_op(Op::I32_OR, line);
    c.emit_if(line);
    get(c, size, line);
    set(c, end, line);
    c.emit_else(line);
    get(c, length_value, line);
    emit_offset_integer(c, line);
    set(c, end, line);
    get(c, end, line);
    c.emit_f64_const(0.0, line);
    c.emit_op(Op::F64_LT, line);
    c.emit_if(line);
    get(c, size, line);
    get(c, end, line);
    c.emit_op(Op::F64_ADD, line);
    set(c, end, line);
    c.emit_else(line);
    get(c, start, line);
    get(c, end, line);
    c.emit_op(Op::F64_ADD, line);
    set(c, end, line);
    c.emit_end(line);
    c.emit_end(line);
    emit_clamp(c, end, start, size, line);

    get(c, text, line);
    import(c, "wasm:js-string", "length", 1, line);
    set(c, units, line);

    // ASCII: byte offsets are code-unit offsets.
    get(c, units, line);
    c.emit_op(Op::F64_FROM_I32, line);
    get(c, size, line);
    c.emit_op(Op::F64_EQ, line);
    c.emit_if(line);
    get(c, text, line);
    get(c, start, line);
    c.emit_op(Op::I32_FROM_F64, line);
    get(c, end, line);
    c.emit_op(Op::I32_FROM_F64, line);
    import(c, "wasm:js-string", "substring", 3, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);

    // Walk code points: `begin`/`finish` are the code units of the first and
    // one-past-last whole code points inside [start, end).
    c.emit_i32_const(0, line);
    set(c, unit, line);
    c.emit_f64_const(0.0, line);
    set(c, byte, line);
    c.emit_i32_const(-1, line);
    set(c, begin, line);
    c.emit_i32_const(-1, line);
    set(c, finish, line);
    c.emit_f64_const(0.0, line);
    set(c, lead, line);
    c.emit_i32_const(0, line);
    set(c, tail, line);
    let block = c.emit_block(line);
    let (lp, _) = c.emit_loop_s(line);
    get(c, unit, line);
    get(c, units, line);
    c.emit_op(Op::I32_GE_S, line);
    c.emit_br_if(1, line);
    // width and code-unit span of the code point at `unit`
    get(c, text, line);
    get(c, unit, line);
    import(c, "wasm:js-string", "codePointAt", 2, line);
    let code = width;
    set(c, code, line);
    get(c, code, line);
    c.emit_i32_const(0xFFFF, line);
    c.emit_op(Op::I32_GT_U, line);
    c.emit_if_i32(line);
    c.emit_i32_const(2, line);
    c.emit_else(line);
    c.emit_i32_const(1, line);
    c.emit_end(line);
    set(c, span, line);
    get(c, code, line);
    c.emit_i32_const(0x80, line);
    c.emit_op(Op::I32_LT_U, line);
    c.emit_if_i32(line);
    c.emit_i32_const(1, line);
    c.emit_else(line);
    get(c, code, line);
    c.emit_i32_const(0x800, line);
    c.emit_op(Op::I32_LT_U, line);
    c.emit_if_i32(line);
    c.emit_i32_const(2, line);
    c.emit_else(line);
    get(c, code, line);
    c.emit_i32_const(0x10000, line);
    c.emit_op(Op::I32_LT_U, line);
    c.emit_if_i32(line);
    c.emit_i32_const(3, line);
    c.emit_else(line);
    c.emit_i32_const(4, line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_op(Op::F64_FROM_I32, line);
    get(c, byte, line);
    c.emit_op(Op::F64_ADD, line);
    set(c, cp_end, line);

    // Not begun yet: does this code point start at/after `start`, or straddle it?
    get(c, begin, line);
    c.emit_i32_const(0, line);
    c.emit_op(Op::I32_LT_S, line);
    c.emit_if(line);
    get(c, byte, line);
    get(c, start, line);
    c.emit_op(Op::F64_GE, line);
    c.emit_if(line);
    get(c, unit, line);
    set(c, begin, line);
    c.emit_else(line);
    get(c, cp_end, line);
    get(c, start, line);
    c.emit_op(Op::F64_GT, line);
    c.emit_if(line);
    // `start` falls inside: its remaining continuation bytes, up to `end`,
    // are undecodable one by one.
    get(c, cp_end, line);
    get(c, end, line);
    c.emit_op(Op::F64_MIN, line);
    get(c, start, line);
    c.emit_op(Op::F64_SUB, line);
    set(c, lead, line);
    get(c, end, line);
    get(c, cp_end, line);
    c.emit_op(Op::F64_LT, line);
    c.emit_if(line);
    c.emit_string_const("", line);
    set(c, out, line);
    emit_append_replacements(c, out, lead, line);
    get(c, out, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);
    get(c, unit, line);
    get(c, span, line);
    c.emit_op(Op::I32_ADD, line);
    set(c, begin, line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_end(line);

    // Begun: does this code point end the slice?
    get(c, begin, line);
    c.emit_i32_const(0, line);
    c.emit_op(Op::I32_GE_S, line);
    get(c, unit, line);
    get(c, begin, line);
    c.emit_op(Op::I32_GE_S, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_if(line);
    get(c, byte, line);
    get(c, end, line);
    c.emit_op(Op::F64_GE, line);
    c.emit_if(line);
    get(c, unit, line);
    set(c, finish, line);
    c.emit_br(3, line);
    c.emit_end(line);
    get(c, cp_end, line);
    get(c, end, line);
    c.emit_op(Op::F64_GT, line);
    c.emit_if(line);
    get(c, unit, line);
    set(c, finish, line);
    c.emit_i32_const(1, line);
    set(c, tail, line);
    c.emit_br(3, line);
    c.emit_end(line);
    c.emit_end(line);

    get(c, cp_end, line);
    set(c, byte, line);
    get(c, unit, line);
    get(c, span, line);
    c.emit_op(Op::I32_ADD, line);
    set(c, unit, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(lp);
    c.emit_end(line);
    c.patch_block(block);

    for slot in [begin, finish] {
        get(c, slot, line);
        c.emit_i32_const(0, line);
        c.emit_op(Op::I32_LT_S, line);
        c.emit_if(line);
        get(c, units, line);
        set(c, slot, line);
        c.emit_end(line);
    }
    c.emit_string_const("", line);
    set(c, out, line);
    emit_append_replacements(c, out, lead, line);
    get(c, out, line);
    get(c, text, line);
    get(c, begin, line);
    get(c, finish, line);
    import(c, "wasm:js-string", "substring", 3, line);
    import(c, "wasm:js-string", "concat", 2, line);
    set(c, out, line);
    get(c, tail, line);
    c.emit_if(line);
    get(c, out, line);
    c.emit_string_const(REPLACEMENT, line);
    import(c, "wasm:js-string", "concat", 2, line);
    set(c, out, line);
    c.emit_end(line);
    get(c, out, line);
    c.emit_op(Op::RETURN, line);
    index
}

/// php `substr($string, $offset, $length = null)` in UTF-8 bytes.
/// Stack: `[string, offset, length?] -> [string]`, the string already
/// converted.
pub fn emit_substr_bytes(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let c = &mut chunks[current];
    if argc < 3 {
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
    let arguments = c.alloc_scratch(3);
    for offset in (0..3).rev() {
        set(c, arguments + offset, line);
    }
    let helper = ensure_substr_helper(chunks);
    let c = &mut chunks[current];
    for offset in 0..3 {
        get(c, arguments + offset, line);
    }
    if vybe_compiler::primitives::functions::try_emit_direct_helper_call(chunks, current, helper, 3, line) {
        return;
    }
    let c = &mut chunks[current];
    for offset in (0..3).rev() {
        set(c, arguments + offset, line);
    }
    c.emit_op_u16(Op::REF_FUNC, helper as u16, line);
    c.emit(0u8, line);
    for offset in 0..3 {
        get(c, arguments + offset, line);
    }
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(c, 3, line);
}
