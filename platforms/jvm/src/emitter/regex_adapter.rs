//! JVM `java.util.regex` adapters.
//!
//! The storage is ECMA RegExp, but the exposed objects are Java/Kotlin-shaped:
//! Pattern carries the compiled regexp and source, Matcher carries input plus
//! the last match state.

use std::sync::Arc;

use vybe_compiler::primitives::{
    collections,
    instructions::host,
    ops,
};
use vybe_runtime::opcode::heaptype;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

const RE_KEY: &str = "__jvm_regex_re";
const SOURCE_KEY: &str = "__jvm_regex_source";
const FLAGS_KEY: &str = "__jvm_regex_flags";
const INPUT_KEY: &str = "__jvm_regex_input";
const MATCH_KEY: &str = "__jvm_regex_match";
const CURSOR_KEY: &str = "__jvm_regex_cursor";
const START_KEY: &str = "__jvm_regex_start";
const END_KEY: &str = "__jvm_regex_end";
const VALUE_KEY: &str = "value";
const RANGE_KEY: &str = "range";
const GROUP_VALUES_KEY: &str = "groupValues";
const DESTRUCTURED_KEY: &str = "destructured";

fn key(chunk: &mut Chunk, name: &str) -> u16 {
    chunk.add_constant(Value::String(Arc::from(name)))
}

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn field_get(chunk: &mut Chunk, slot: u16, field: &str, line: u32) {
    vybe_compiler::primitives::class_slots::emit_class_get(
        chunk,
        vybe_compiler::primitives::class_slots::ObjSource::Local(slot),
        &super::object_fields::field_slot(field),
        vybe_compiler::primitives::class_slots::Dest::Stack,
        line,
    );
}

fn field_set_from_stack(chunk: &mut Chunk, slot: u16, field: &str, line: u32) {
    let value = chunk.alloc_scratch(1);
    set(chunk, value, line);
    vybe_compiler::primitives::class_slots::emit_class_set(
        chunk,
        vybe_compiler::primitives::class_slots::ObjSource::Local(slot),
        &super::object_fields::field_slot(field),
        vybe_compiler::primitives::class_slots::ValueSource::Local(value),
        line,
    );
}

fn array_get_const(chunk: &mut Chunk, slot: u16, idx: f64, line: u32) {
    get(chunk, slot, line);
    chunk.emit_f64_const(idx, line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

fn object_get_const(chunk: &mut Chunk, slot: u16, field: &str, line: u32) {
    get(chunk, slot, line);
    chunk.emit_string_const(field, line);
    host::emit(chunk, "ecma:object", "get", 2, line);
}

fn null(chunk: &mut Chunk, line: u32) {
    chunk.emit_ref_null(heaptype::HT_EXTERN, line);
}

fn emit_clone_matcher_from_slot(chunks: &mut [Chunk], current: usize, matcher: u16, line: u32) {
    vybe_compiler::primitives::class_slots::emit_class_alloc(&mut chunks[current], line);
    let copy = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], copy, line);
    for field in [RE_KEY, INPUT_KEY, MATCH_KEY, CURSOR_KEY, START_KEY, END_KEY] {
        field_get(&mut chunks[current], matcher, field, line);
        field_set_from_stack(&mut chunks[current], copy, field, line);
    }
    field_get(&mut chunks[current], matcher, MATCH_KEY, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    field_set_from_stack(&mut chunks[current], copy, VALUE_KEY, line);
    field_get(&mut chunks[current], matcher, MATCH_KEY, line);
    field_set_from_stack(&mut chunks[current], copy, GROUP_VALUES_KEY, line);
    field_get(&mut chunks[current], matcher, MATCH_KEY, line);
    chunks[current].emit_f64_const(1.0, line);
    host::emit(&mut chunks[current], "ecma:array", "slice", 2, line);
    field_set_from_stack(&mut chunks[current], copy, DESTRUCTURED_KEY, line);
    emit_match_range_object_from_slot(chunks, current, matcher, line);
    field_set_from_stack(&mut chunks[current], copy, RANGE_KEY, line);
    get(&mut chunks[current], copy, line);
}

fn emit_match_range_object_from_slot(chunks: &mut [Chunk], current: usize, matcher: u16, line: u32) {
    let chunk = &mut chunks[current];
    vybe_compiler::primitives::class_slots::emit_class_alloc(chunk, line);
    let range = chunk.alloc_scratch(1);
    set(chunk, range, line);
    field_get(chunk, matcher, START_KEY, line);
    field_set_from_stack(chunk, range, "first", line);
    field_get(chunk, matcher, START_KEY, line);
    field_set_from_stack(chunk, range, "start", line);
    field_get(chunk, matcher, END_KEY, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_SUB, line);
    let last = chunk.alloc_scratch(1);
    set(chunk, last, line);
    get(chunk, last, line);
    field_set_from_stack(chunk, range, "last", line);
    get(chunk, last, line);
    field_set_from_stack(chunk, range, "endInclusive", line);
    get(chunk, range, line);
}

pub fn emit_pattern_compile(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc >= 2 {
        emit_pattern_compile_flags(chunks, current, argc, line);
        return;
    }
    let chunk = &mut chunks[current];
    let source = chunk.alloc_scratch(1);
    set(chunk, source, line);

    get(chunk, source, line);
    host::emit(chunk, "ecma:regexp", "new", 1, line);
    let re = chunk.alloc_scratch(1);
    set(chunk, re, line);

    vybe_compiler::primitives::class_slots::emit_class_alloc(chunk, line);
    let pat = chunk.alloc_scratch(1);
    set(chunk, pat, line);
    get(chunk, source, line);
    field_set_from_stack(chunk, pat, SOURCE_KEY, line);
    get(chunk, re, line);
    field_set_from_stack(chunk, pat, RE_KEY, line);
    chunk.emit_f64_const(0.0, line);
    field_set_from_stack(chunk, pat, FLAGS_KEY, line);
    get(chunk, pat, line);
}

pub fn emit_pattern_compile_flags(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 2..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let flags = chunk.alloc_scratch(1);
    let source = chunk.alloc_scratch(1);
    set(chunk, flags, line);
    set(chunk, source, line);

    get(chunk, source, line);
    // Flag translation is intentionally layered: the object keeps Java's
    // numeric flags now; option-specific ECMA flags can widen here without
    // changing the namespace surface.
    host::emit(chunk, "ecma:regexp", "new", 1, line);
    let re = chunk.alloc_scratch(1);
    set(chunk, re, line);

    vybe_compiler::primitives::class_slots::emit_class_alloc(chunk, line);
    let pat = chunk.alloc_scratch(1);
    set(chunk, pat, line);
    get(chunk, source, line);
    field_set_from_stack(chunk, pat, SOURCE_KEY, line);
    get(chunk, re, line);
    field_set_from_stack(chunk, pat, RE_KEY, line);
    get(chunk, flags, line);
    field_set_from_stack(chunk, pat, FLAGS_KEY, line);
    get(chunk, pat, line);
}

pub fn emit_pattern_pattern(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let pattern = chunk.alloc_scratch(1);
    set(chunk, pattern, line);
    field_get(chunk, pattern, SOURCE_KEY, line);
}

pub fn emit_pattern_flags(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let pattern = chunk.alloc_scratch(1);
    set(chunk, pattern, line);
    field_get(chunk, pattern, FLAGS_KEY, line);
}

pub fn emit_pattern_matcher(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let input = chunk.alloc_scratch(1);
    let pattern = chunk.alloc_scratch(1);
    set(chunk, input, line);
    set(chunk, pattern, line);
    vybe_compiler::primitives::class_slots::emit_class_alloc(chunk, line);
    let matcher = chunk.alloc_scratch(1);
    set(chunk, matcher, line);
    field_get(chunk, pattern, RE_KEY, line);
    field_set_from_stack(chunk, matcher, RE_KEY, line);
    get(chunk, input, line);
    field_set_from_stack(chunk, matcher, INPUT_KEY, line);
    null(chunk, line);
    field_set_from_stack(chunk, matcher, MATCH_KEY, line);
    chunk.emit_f64_const(0.0, line);
    field_set_from_stack(chunk, matcher, CURSOR_KEY, line);
    chunk.emit_f64_const(-1.0, line);
    field_set_from_stack(chunk, matcher, START_KEY, line);
    chunk.emit_f64_const(-1.0, line);
    field_set_from_stack(chunk, matcher, END_KEY, line);
    get(chunk, matcher, line);
}

pub fn emit_matcher_reset(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let input = (argc >= 2).then(|| chunk.alloc_scratch(1));
    let matcher = chunk.alloc_scratch(1);
    if let Some(slot) = input {
        set(chunk, slot, line);
    }
    set(chunk, matcher, line);
    if let Some(slot) = input {
        get(chunk, slot, line);
        field_set_from_stack(chunk, matcher, INPUT_KEY, line);
    }
    null(chunk, line);
    field_set_from_stack(chunk, matcher, MATCH_KEY, line);
    chunk.emit_f64_const(0.0, line);
    field_set_from_stack(chunk, matcher, CURSOR_KEY, line);
    get(chunk, matcher, line);
}

pub fn emit_matcher_find(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let matcher = chunk.alloc_scratch(1);
    set(chunk, matcher, line);
    field_get(chunk, matcher, INPUT_KEY, line);
    field_get(chunk, matcher, CURSOR_KEY, line);
    host::emit(chunk, "ecma:string", "slice", 2, line);
    let haystack = chunk.alloc_scratch(1);
    set(chunk, haystack, line);
    field_get(chunk, matcher, RE_KEY, line);
    get(chunk, haystack, line);
    host::emit(chunk, "ecma:regexp", "exec", 2, line);
    let result = chunk.alloc_scratch(1);
    set(chunk, result, line);
    get(chunk, result, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    null(chunk, line);
    field_set_from_stack(chunk, matcher, MATCH_KEY, line);
    chunk.emit_bool_const(false, line);
    chunk.emit_else(line);
    get(chunk, result, line);
    field_set_from_stack(chunk, matcher, MATCH_KEY, line);
    field_get(chunk, matcher, CURSOR_KEY, line);
    host::emit(chunk, "wasm:js-number", "toF64", 1, line);
    object_get_const(chunk, result, "index", line);
    host::emit(chunk, "wasm:js-number", "toF64", 1, line);
    chunk.emit_op(Op::F64_ADD, line);
    let start = chunk.alloc_scratch(1);
    set(chunk, start, line);
    get(chunk, start, line);
    field_set_from_stack(chunk, matcher, START_KEY, line);
    get(chunk, start, line);
    array_get_const(chunk, result, 0.0, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    chunk.emit_op(Op::F64_FROM_I32, line);
    chunk.emit_op(Op::F64_ADD, line);
    let end = chunk.alloc_scratch(1);
    set(chunk, end, line);
    get(chunk, end, line);
    field_set_from_stack(chunk, matcher, END_KEY, line);
    get(chunk, end, line);
    get(chunk, start, line);
    ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    get(chunk, end, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_else(line);
    get(chunk, end, line);
    chunk.emit_end(line);
    field_set_from_stack(chunk, matcher, CURSOR_KEY, line);
    field_get(chunk, matcher, MATCH_KEY, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    field_set_from_stack(chunk, matcher, VALUE_KEY, line);
    field_get(chunk, matcher, MATCH_KEY, line);
    field_set_from_stack(chunk, matcher, GROUP_VALUES_KEY, line);
    field_get(chunk, matcher, MATCH_KEY, line);
    chunk.emit_f64_const(1.0, line);
    host::emit(chunk, "ecma:array", "slice", 2, line);
    field_set_from_stack(chunk, matcher, DESTRUCTURED_KEY, line);
    vybe_compiler::primitives::class_slots::emit_class_alloc(chunk, line);
    let range = chunk.alloc_scratch(1);
    set(chunk, range, line);
    field_get(chunk, matcher, START_KEY, line);
    field_set_from_stack(chunk, range, "first", line);
    field_get(chunk, matcher, START_KEY, line);
    field_set_from_stack(chunk, range, "start", line);
    field_get(chunk, matcher, END_KEY, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_SUB, line);
    let last = chunk.alloc_scratch(1);
    set(chunk, last, line);
    get(chunk, last, line);
    field_set_from_stack(chunk, range, "last", line);
    get(chunk, last, line);
    field_set_from_stack(chunk, range, "endInclusive", line);
    get(chunk, range, line);
    field_set_from_stack(chunk, matcher, RANGE_KEY, line);
    chunk.emit_bool_const(true, line);
    chunk.emit_end(line);
}

pub fn emit_matcher_matches(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let matcher = chunk.alloc_scratch(1);
    set(chunk, matcher, line);
    field_get(chunk, matcher, RE_KEY, line);
    field_get(chunk, matcher, INPUT_KEY, line);
    host::emit(chunk, "ecma:regexp", "exec", 2, line);
    let result = chunk.alloc_scratch(1);
    set(chunk, result, line);
    get(chunk, result, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_else(line);
    object_get_const(chunk, result, "index", line);
    chunk.emit_f64_const(0.0, line);
    ops::emit_dyn_eq(chunk, line);
    array_get_const(chunk, result, 0.0, line);
    field_get(chunk, matcher, INPUT_KEY, line);
    ops::emit_dyn_eq(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    ops::emit_i32_to_bool(chunk, line);
    chunk.emit_end(line);
}

pub fn emit_matcher_group(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let idx = if argc >= 2 {
        let idx = chunk.alloc_scratch(1);
        set(chunk, idx, line);
        Some(idx)
    } else {
        None
    };
    let matcher = chunk.alloc_scratch(1);
    set(chunk, matcher, line);
    field_get(chunk, matcher, MATCH_KEY, line);
    if let Some(idx) = idx {
        get(chunk, idx, line);
    } else {
        chunk.emit_f64_const(0.0, line);
    }
    chunk.emit_op(Op::ARRAY_GET, line);
}

pub fn emit_matcher_start(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let matcher = chunk.alloc_scratch(1);
    set(chunk, matcher, line);
    field_get(chunk, matcher, START_KEY, line);
}

pub fn emit_matcher_end(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let matcher = chunk.alloc_scratch(1);
    set(chunk, matcher, line);
    field_get(chunk, matcher, END_KEY, line);
}

pub fn emit_pattern_split(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let limit = (argc >= 3).then(|| chunk.alloc_scratch(1));
    let input = chunk.alloc_scratch(1);
    let pattern = chunk.alloc_scratch(1);
    if let Some(slot) = limit {
        set(chunk, slot, line);
    }
    set(chunk, input, line);
    set(chunk, pattern, line);
    get(chunk, input, line);
    field_get(chunk, pattern, RE_KEY, line);
    if let Some(slot) = limit {
        get(chunk, slot, line);
        host::emit(chunk, "ecma:regexp", "split", 3, line);
    } else {
        host::emit(chunk, "ecma:regexp", "split", 2, line);
    }
}

pub fn emit_pattern_replace_all(chunks: &mut [Chunk], current: usize, first_only: bool, line: u32) {
    let chunk = &mut chunks[current];
    let replacement = chunk.alloc_scratch(1);
    let input = chunk.alloc_scratch(1);
    let pattern = chunk.alloc_scratch(1);
    set(chunk, replacement, line);
    set(chunk, input, line);
    set(chunk, pattern, line);
    get(chunk, input, line);
    field_get(chunk, pattern, RE_KEY, line);
    get(chunk, replacement, line);
    host::emit(
        chunk,
        "ecma:regexp",
        if first_only { "replace" } else { "replaceAll" },
        3,
        line,
    );
}

pub fn emit_pattern_match_full(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_pattern_matcher(chunks, current, line);
    emit_matcher_matches(chunks, current, line);
}

pub fn emit_pattern_match_entire(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_pattern_matcher(chunks, current, line);
    let matcher = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], matcher, line);
    get(&mut chunks[current], matcher, line);
    emit_matcher_find(chunks, current, line);
    chunks[current].emit_if_value(line);
    field_get(&mut chunks[current], matcher, START_KEY, line);
    chunks[current].emit_f64_const(0.0, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    field_get(&mut chunks[current], matcher, VALUE_KEY, line);
    field_get(&mut chunks[current], matcher, INPUT_KEY, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_AND, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], matcher, line);
    chunks[current].emit_else(line);
    null(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    null(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

pub fn emit_pattern_contains(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let input = chunk.alloc_scratch(1);
    let pattern = chunk.alloc_scratch(1);
    set(chunk, input, line);
    set(chunk, pattern, line);
    field_get(chunk, pattern, RE_KEY, line);
    get(chunk, input, line);
    host::emit(chunk, "ecma:regexp", "test", 2, line);
}

pub fn emit_pattern_find(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let start = if argc >= 3 {
        let slot = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], slot, line);
        Some(slot)
    } else {
        None
    };
    emit_pattern_matcher(chunks, current, line);
    let matcher = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], matcher, line);
    if let Some(start) = start {
        get(&mut chunks[current], start, line);
        field_set_from_stack(&mut chunks[current], matcher, CURSOR_KEY, line);
    }
    get(&mut chunks[current], matcher, line);
    emit_matcher_find(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], matcher, line);
    chunks[current].emit_else(line);
    null(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

pub fn emit_pattern_matches_at(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let start = if argc >= 3 {
        let slot = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], slot, line);
        slot
    } else {
        chunks[current].emit_f64_const(0.0, line);
        let slot = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], slot, line);
        slot
    };
    emit_pattern_matcher(chunks, current, line);
    let matcher = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], matcher, line);
    get(&mut chunks[current], start, line);
    field_set_from_stack(&mut chunks[current], matcher, CURSOR_KEY, line);
    get(&mut chunks[current], matcher, line);
    emit_matcher_find(chunks, current, line);
    chunks[current].emit_if_value(line);
    field_get(&mut chunks[current], matcher, START_KEY, line);
    get(&mut chunks[current], start, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

pub fn emit_pattern_find_all(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let start = if argc >= 3 {
        let slot = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], slot, line);
        Some(slot)
    } else {
        None
    };
    emit_pattern_matcher(chunks, current, line);
    let matcher = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], matcher, line);
    if let Some(start) = start {
        get(&mut chunks[current], start, line);
        field_set_from_stack(&mut chunks[current], matcher, CURSOR_KEY, line);
    }
    collections::emit_array_new(chunks, current, 0, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);

    let block = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], matcher, line);
    emit_matcher_find(chunks, current, line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], out, line);
    emit_clone_matcher_from_slot(chunks, current, matcher, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    get(&mut chunks[current], out, line);
}

pub fn emit_match_result_value(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let matcher = chunk.alloc_scratch(1);
    set(chunk, matcher, line);
    field_get(chunk, matcher, MATCH_KEY, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
}

pub fn emit_match_result_range(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let matcher = chunk.alloc_scratch(1);
    set(chunk, matcher, line);
    emit_match_range_object_from_slot(chunks, current, matcher, line);
}

pub fn emit_to_pattern(chunks: &mut [Chunk], current: usize, _line: u32) {
    // Kotlin Regex is represented by the same Pattern object.
    let _ = current;
    let _ = chunks;
}
