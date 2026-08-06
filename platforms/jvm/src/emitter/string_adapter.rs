//! JVM `java.lang.String` and `java.util.Objects` adapters.

use vybe_compiler::primitives::collections;
use vybe_compiler::primitives::strings;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

pub fn emit_join(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let elem_count = argc.saturating_sub(1);
    let first_elem = chunks[current].alloc_scratch(elem_count as u16);
    let delimiter = chunks[current].alloc_scratch(1);
    for k in (0..elem_count).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, first_elem + k as u16, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, delimiter, line);

    if elem_count == 1 {
        let elem = first_elem;
        chunks[current].emit_op_u16(Op::LOCAL_GET, elem, line);
        let len = chunks[current].add_import("ecma:array", "length");
        chunks[current].emit_call(len, 1, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if(line);
        collections::emit_array_new(chunks, current, 0, line);
        let array = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_SET, array, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, array, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, elem, line);
        let push = chunks[current].add_import("ecma:array", "push");
        chunks[current].emit_call(push, 2, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, array, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, elem, line);
        chunks[current].emit_end(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, delimiter, line);
        let join = chunks[current].add_import("ecma:array", "join");
        chunks[current].emit_call(join, 2, line);
        return;
    }

    collections::emit_array_new(chunks, current, 0, line);
    let array = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, array, line);
    let push = chunks[current].add_import("ecma:array", "push");
    for k in 0..elem_count {
        chunks[current].emit_op_u16(Op::LOCAL_GET, array, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, first_elem + k as u16, line);
        chunks[current].emit_call(push, 2, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, array, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, delimiter, line);
    let join = chunks[current].add_import("ecma:array", "join");
    chunks[current].emit_call(join, 2, line);
}

pub fn emit_require_non_null(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc > 1 {
        chunks[current].emit_op(Op::DROP, line);
    }
}

// ── java.lang.Character classifiers ─────────────────────────────────────
//
// The classification itself is the SHARED tier-3 string primitive
// (`vybe_compiler::primitives::strings::emit_is_*` — the targets of the
// `(String, Is*)` platform slot rows in builtin_slots.rs). These wrappers
// own only what is JVM: the char model admits non-strings — a char VALUE
// is a one-char string, but a lone surrogate is a NUMBER (see the java
// walker's char-literal rule) — and a non-string char is in no class, so
// the guard answers `false` before the primitive (which requires a string
// receiver) runs. Same split as php coercing before `emit_byte_length`.

/// Stack: `[c]` → `[bool]` — string-test guard, then the shared class
/// primitive; a number-model char is `false` in every class.
fn emit_char_classified(
    chunks: &mut [Chunk],
    current: usize,
    class: fn(&mut [Chunk], usize, u32),
    line: u32,
) {
    let v = chunks[current].alloc_scratch(1);
    {
        let c = &mut chunks[current];
        c.emit_op_u16(Op::LOCAL_SET, v, line);
        c.emit_op_u16(Op::LOCAL_GET, v, line);
        let test = c.add_import("wasm:js-string", "test");
        c.emit_call(test, 1, line);
        c.emit_if_value(line);
        c.emit_op_u16(Op::LOCAL_GET, v, line);
    }
    class(chunks, current, line);
    let c = &mut chunks[current];
    c.emit_else(line);
    c.emit_bool_const(false, line);
    c.emit_end(line);
}

pub fn emit_char_is_upper(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_char_classified(chunks, current, strings::emit_is_upper, line);
}

pub fn emit_char_is_lower(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_char_classified(chunks, current, strings::emit_is_lower, line);
}

pub fn emit_char_is_letter(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_char_classified(chunks, current, strings::emit_is_alpha, line);
}

pub fn emit_char_is_digit(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_char_classified(chunks, current, strings::emit_is_digit, line);
}

pub fn emit_char_is_alnum(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_char_classified(chunks, current, strings::emit_is_alnum, line);
}

pub fn emit_char_is_space(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_char_classified(chunks, current, strings::emit_is_space, line);
}
