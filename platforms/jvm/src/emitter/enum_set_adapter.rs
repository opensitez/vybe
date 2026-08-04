//! `java.util.EnumSet` — a JDK collection, so it lives with the rest of the
//! JDK and every JVM language reaches the same one.
//!
//! Backed by an array of enum display NAMES, with the enum's full
//! declaration-ordered name list attached as `__java_enum_names` so an
//! ordinal can be resolved either way. It used to sit in `languages/java`,
//! which made a `java.util` class the property of one frontend — Kotlin would
//! have had to declare all fifteen operations again.
//!
//! ## Every operation derives its data from its own arguments
//!
//! These are tree LEAVES (`jvm.java.util.EnumSet.*`), and a leaf receives only
//! what the source wrote. Earlier the Java walker prepended a compile-time
//! name array to every call and folded each constant to its ordinal, so the
//! adapters read `names[ordinal]` — which is why they were unreachable from
//! any frontend but Java. Now:
//!
//! | operation | where the names come from |
//! |---|---|
//! | `of`, `range` | the constant argument's [`NAMES_KEY`] |
//! | `noneOf`, `allOf` | the registry, keyed by the `X.class` name string |
//! | `copyOf`, `complementOf`, every instance method | the set's own [`NAMES_KEY`] |
//!
//! and a value's display name is `value.__name` — read off the constant, which
//! is a `java.lang.Enum` instance that carries it (`lang_enum`).

use super::enum_adapter;
use vybe_compiler::primitives::{
    collections,
    instructions::{core_wasm, host} };
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// Shared with `lang_enum`, which stamps the same key on every constant — one
/// spelling, so the two sides cannot drift apart.
const NAMES_KEY: &str = crate::lang_enum::NAMES_FIELD;
const CLASS_KEY: &str = crate::lang_enum::CLASS_FIELD;

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn attach_names(chunks: &mut [Chunk], current: usize, set_slot: u16, names_slot: u16, line: u32) {
    get(&mut chunks[current], set_slot, line);
    chunks[current].emit_string_const(NAMES_KEY, line);
    get(&mut chunks[current], names_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], set_slot, line);
    chunks[current].emit_string_const(CLASS_KEY, line);
    chunks[current].emit_string_const("EnumSet", line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// Read property `key` off the object in `slot`.
fn emit_member(chunks: &mut [Chunk], current: usize, slot: u16, key: &str, line: u32) {
    get(&mut chunks[current], slot, line);
    chunks[current].emit_string_const(key, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
}

fn emit_names(chunks: &mut [Chunk], current: usize, set_slot: u16, line: u32) {
    emit_member(chunks, current, set_slot, NAMES_KEY, line);
}

/// A constant's display name. It is `value.__name` — the field
/// `java.lang.Enum`'s constructor stamps — not an index into the set's name
/// list, which only worked while the Java walker was constant-folding
/// `Color.RED` to its ordinal at the call site.
fn emit_name_for_value(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) {
    emit_member(chunks, current, value_slot, crate::lang_enum::NAME_FIELD, line);
}

fn emit_contains_name(
    chunks: &mut [Chunk],
    current: usize,
    set_slot: u16,
    name_slot: u16,
    line: u32,
) {
    get(&mut chunks[current], set_slot, line);
    get(&mut chunks[current], name_slot, line);
    collections::emit_contains(chunks, current, line);
}

fn push_name_if_absent(
    chunks: &mut [Chunk],
    current: usize,
    set_slot: u16,
    name_slot: u16,
    line: u32,
) {
    emit_contains_name(chunks, current, set_slot, name_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], set_slot, line);
    get(&mut chunks[current], name_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
}

/// `EnumSet.noneOf(Color.class)` — stack `[className]`.
pub fn emit_none_of(chunks: &mut [Chunk], current: usize, line: u32) {
    let names = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    enum_adapter::emit_constants_of(chunks, current, line);
    set(&mut chunks[current], names, line);
    collections::emit_array_new(chunks, current, 0, line);
    set(&mut chunks[current], out, line);
    attach_names(chunks, current, out, names, line);
    get(&mut chunks[current], out, line);
}

/// `EnumSet.allOf(Color.class)` — stack `[className]`.
pub fn emit_all_of(chunks: &mut [Chunk], current: usize, line: u32) {
    let names = chunks[current].alloc_scratch(1);
    enum_adapter::emit_constants_of(chunks, current, line);
    set(&mut chunks[current], names, line);
    get(&mut chunks[current], names, line);
    collections::emit_clone(chunks, current, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    attach_names(chunks, current, out, names, line);
    get(&mut chunks[current], out, line);
}

/// `EnumSet.of(RED, BLUE, …)` — stack `[v1 … vn]`, all of them constants.
///
/// The enum's full name list comes off the FIRST argument: a constant knows
/// its own declaring enum, so the call site does not have to say it twice.
pub fn emit_of(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let values = chunks[current].alloc_scratch(argc.max(1) as u16);
    for i in (0..argc).rev() {
        set(&mut chunks[current], values + i as u16, line);
    }
    let names = chunks[current].alloc_scratch(1);
    emit_member(chunks, current, values, NAMES_KEY, line);
    set(&mut chunks[current], names, line);
    collections::emit_array_new(chunks, current, 0, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    attach_names(chunks, current, out, names, line);
    for i in 0..argc {
        emit_name_for_value(chunks, current, values + i as u16, line);
        let name = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], name, line);
        push_name_if_absent(chunks, current, out, name, line);
    }
    get(&mut chunks[current], out, line);
}

pub fn emit_copy_of(chunks: &mut [Chunk], current: usize, line: u32) {
    let source = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], source, line);
    get(&mut chunks[current], source, line);
    collections::emit_clone(chunks, current, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    get(&mut chunks[current], source, line);
    chunks[current].emit_string_const(NAMES_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    let names = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], names, line);
    attach_names(chunks, current, out, names, line);
    get(&mut chunks[current], out, line);
}

pub fn emit_complement_of(chunks: &mut [Chunk], current: usize, line: u32) {
    let source = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], source, line);
    emit_names(chunks, current, source, line);
    let names = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], names, line);
    collections::emit_array_new(chunks, current, 0, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    attach_names(chunks, current, out, names, line);

    let i = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let name = chunks[current].alloc_scratch(1);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], i, line);
    get(&mut chunks[current], names, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    let _block = chunks[current].emit_block(line);
    let (_loop, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], len, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], names, line);
    get(&mut chunks[current], i, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], name, line);
    emit_contains_name(chunks, current, source, name, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], name, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], out, line);
}

/// `EnumSet.range(FROM, TO)` — stack `[from, to]`, inclusive on both ends.
pub fn emit_range(chunks: &mut [Chunk], current: usize, line: u32) {
    let end = chunks[current].alloc_scratch(1);
    let start = chunks[current].alloc_scratch(1);
    let names = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], end, line);
    set(&mut chunks[current], start, line);
    emit_member(chunks, current, start, NAMES_KEY, line);
    set(&mut chunks[current], names, line);
    // The bounds are CONSTANTS; their ordinals are the array indices. Each
    // read completes before its own slot is overwritten.
    emit_member(chunks, current, start, crate::lang_enum::ORDINAL_FIELD, line);
    set(&mut chunks[current], start, line);
    emit_member(chunks, current, end, crate::lang_enum::ORDINAL_FIELD, line);
    set(&mut chunks[current], end, line);
    collections::emit_array_new(chunks, current, 0, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    attach_names(chunks, current, out, names, line);
    let i = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], start, line);
    set(&mut chunks[current], i, line);
    let _block = chunks[current].emit_block(line);
    let (_loop, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], end, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], names, line);
    get(&mut chunks[current], i, line);
    collections::emit_get(chunks, current, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], out, line);
}

pub fn emit_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let set_slot = chunks[current].alloc_scratch(1);
    let name = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], set_slot, line);
    emit_name_for_value(chunks, current, value, line);
    set(&mut chunks[current], name, line);
    emit_contains_name(chunks, current, set_slot, name, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], set_slot, line);
    get(&mut chunks[current], name, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_end(line);
}

pub fn emit_add_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let source = chunks[current].alloc_scratch(1);
    let target = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let name = chunks[current].alloc_scratch(1);
    let changed = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], source, line);
    set(&mut chunks[current], target, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], i, line);
    get(&mut chunks[current], source, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], changed, line);
    let _block = chunks[current].emit_block(line);
    let (_loop, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], len, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], source, line);
    get(&mut chunks[current], i, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], name, line);
    emit_contains_name(chunks, current, target, name, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], target, line);
    get(&mut chunks[current], name, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);
    set(&mut chunks[current], changed, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], changed, line);
}

pub fn emit_contains(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let set_slot = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], set_slot, line);
    emit_name_for_value(chunks, current, value, line);
    let name = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], name, line);
    emit_contains_name(chunks, current, set_slot, name, line);
}

pub fn emit_contains_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let source = chunks[current].alloc_scratch(1);
    let target = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let name = chunks[current].alloc_scratch(1);
    let ok = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], source, line);
    set(&mut chunks[current], target, line);
    chunks[current].emit_bool_const(true, line);
    set(&mut chunks[current], ok, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], i, line);
    get(&mut chunks[current], source, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    let _block = chunks[current].emit_block(line);
    let (_loop, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], len, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], source, line);
    get(&mut chunks[current], i, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], name, line);
    emit_contains_name(chunks, current, target, name, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], ok, line);
    get(&mut chunks[current], len, line);
    set(&mut chunks[current], i, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], ok, line);
}

pub fn emit_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    let other = chunks[current].alloc_scratch(1);
    let set_slot = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], other, line);
    set(&mut chunks[current], set_slot, line);
    get(&mut chunks[current], set_slot, line);
    collections::emit_len(chunks, current, line);
    get(&mut chunks[current], other, line);
    collections::emit_len(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], set_slot, line);
    get(&mut chunks[current], other, line);
    emit_contains_all(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

pub fn emit_hash_code(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
}

pub fn emit_iterator(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:array", "values", 1, line);
}

pub fn emit_remove(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let set_slot = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], set_slot, line);
    emit_name_for_value(chunks, current, value, line);
    let name = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], name, line);
    get(&mut chunks[current], set_slot, line);
    get(&mut chunks[current], name, line);
    collections::emit_remove_value(chunks, current, line);
}

pub fn emit_get_class(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_string_const("EnumSet", line);
}
