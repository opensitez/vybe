//! JVM `java.lang.String` and `java.util.Objects` adapters.

use vybe_compiler::primitives::collections;
use vybe_compiler::primitives::instructions::host;
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

// ── Moved from `languages/java` ──────────────────────────────────────────
//
// `java.lang.String`, `Character` and `Integer` behaviour that every JVM
// language shares. The two copies of this file overlapped on exactly ONE
// function, `emit_join`, and differed there only in local variable names —
// identical emitted bytecode — so the platform's was kept and the rest simply
// appended. Nothing was reconciled because nothing conflicted.

pub fn emit_index_of(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    host::emit(&mut chunks[current], "ecma:string", "indexOf", argc, line);
}


/// `(char)n` — `String.fromCharCode`, EXCEPT a lone surrogate stays a
/// NUMBER: the string host cannot hold one (it renders empty), which then
/// breaks every downstream `charCodeAt`. Mirrors the walker's rule for a
/// `'\uD800'` literal, so both spellings of a lone surrogate share one
/// model.
pub fn emit_from_char_code(chunks: &mut [Chunk], current: usize, line: u32) {
    let n = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_f64_const(55296.0, line);
    chunks[current].emit_op(Op::F64_GE, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_f64_const(57344.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    host::emit(&mut chunks[current], "ecma:string", "fromCharCode", 1, line);
    chunks[current].emit_end(line);
}


pub fn emit_last_index_of(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let arg_count = argc.saturating_sub(1);
    let first_arg_slot = chunk.alloc_scratch(arg_count as u16);
    let self_slot = chunk.alloc_scratch(1);
    for k in (0..arg_count).rev() {
        chunk.emit_op_u16(Op::LOCAL_SET, first_arg_slot + k as u16, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, self_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, self_slot, line);
    chunk.emit_string_const("lastIndexOf", line);
    for k in 0..arg_count {
        chunk.emit_op_u16(Op::LOCAL_GET, first_arg_slot + k as u16, line);
    }
    let invoke = chunk.add_import("ecma:value", "invokeMethod");
    chunk.emit_call(invoke, argc.saturating_add(1), line);
}


pub fn emit_starts_with(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    host::emit(
        &mut chunks[current],
        "ecma:string",
        "startsWith",
        argc,
        line,
    );
}


/// `String.valueOf(x)` — `java.lang.String`, so the rendering itself lives in
/// `platforms/jvm` and Kotlin reaches the identical one. This was
/// `ecma:string.String`, which never consults the object's ToString slot.
pub fn emit_value_of(chunks: &mut [Chunk], current: usize, line: u32) {
    crate::emitter::object_adapter::emit_to_string(chunks, current, line);
}


pub fn emit_concat(chunks: &mut [Chunk], current: usize, line: u32) {
    let right_slot = chunks[current].alloc_scratch(1);
    let left_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    emit_value_of(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    emit_value_of(chunks, current, line);
    vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);
}


pub fn emit_compare_to(chunks: &mut [Chunk], current: usize, line: u32) {
    let cmp = chunks[current].add_import("ecma:string", "localeCompare");
    chunks[current].emit_call(cmp, 2, line);
}


pub fn emit_char_ord(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    host::emit(&mut chunks[current], "wasm:js-string", "test", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_i32_const(0, line);
    host::emit(
        &mut chunks[current],
        "wasm:js-string",
        "charCodeAt",
        2,
        line,
    );
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_end(line);
}


pub fn emit_trunc_cast(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    host::emit(&mut chunks[current], "wasm:js-string", "test", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_i32_const(0, line);
    host::emit(
        &mut chunks[current],
        "wasm:js-string",
        "charCodeAt",
        2,
        line,
    );
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    vybe_compiler::primitives::math::emit_trunc(&mut chunks[current], line);
    chunks[current].emit_end(line);
}


pub fn emit_compare_ignore_case(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    lower_pair(chunk, line);
    let cmp = chunk.add_import("ecma:string", "localeCompare");
    chunk.emit_call(cmp, 2, line);
}


pub fn emit_equals_ignore_case(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    lower_pair(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
}


pub fn emit_hash_code(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_i32_const(0, line);
}


pub fn emit_matches(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let pattern_slot = chunk.alloc_scratch(1);
    let self_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, pattern_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, self_slot, line);

    let regex_new = chunk.add_import("ecma:regexp", "new");
    let regex_test = chunk.add_import("ecma:regexp", "test");
    chunk.emit_op_u16(Op::LOCAL_GET, pattern_slot, line);
    chunk.emit_call(regex_new, 1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, self_slot, line);
    chunk.emit_call(regex_test, 2, line);
}


pub fn emit_replace_regex(chunks: &mut [Chunk], current: usize, replace_all: bool, line: u32) {
    let chunk = &mut chunks[current];
    let replacement_slot = chunk.alloc_scratch(1);
    let pattern_slot = chunk.alloc_scratch(1);
    let self_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, replacement_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, pattern_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, self_slot, line);

    let regex_new = chunk.add_import("ecma:regexp", "new");
    let replace_name = if replace_all { "replaceAll" } else { "replace" };
    let replace = chunk.add_import("ecma:regexp", replace_name);
    chunk.emit_op_u16(Op::LOCAL_GET, self_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, pattern_slot, line);
    chunk.emit_call(regex_new, 1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, replacement_slot, line);
    chunk.emit_call(replace, 3, line);
}


pub fn emit_to_char_array(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_string_const("", line);
    let idx = chunks[current].add_import("ecma:string", "split");
    chunks[current].emit_call(idx, 2, line);
}


fn lower_pair(chunk: &mut Chunk, line: u32) {
    let other_slot = chunk.alloc_scratch(1);
    let self_slot = chunk.alloc_scratch(1);
    let lower = chunk.add_import("ecma:string", "toLowerCase");
    chunk.emit_op_u16(Op::LOCAL_SET, other_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, self_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, self_slot, line);
    chunk.emit_call(lower, 1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, other_slot, line);
    chunk.emit_call(lower, 1, line);
}
