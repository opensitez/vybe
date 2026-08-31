//! `System.Text.Json.JsonNamingPolicy` — the four separator policies and camelCase.
//!
//! ## The word-boundary rule, measured against .NET SDK 10
//!
//! ```text
//!   ItemValue        -> item_value
//!   HTTPResponse     -> http_response          not h_t_t_p_response
//!   XMLHttpRequest2  -> xml_http_request2
//!   alreadylower     -> alreadylower
//!   A                -> a
//! ```
//!
//! ⛔ It is NOT "split before every capital". `HTTPResponse` shows an acronym
//! does not split internally; `XMLHttpRequest2` shows it DOES split at the last
//! capital of a run when a lowercase follows. So a separator goes before an
//! uppercase at index `i` when `i > 0` AND (the previous character is not
//! uppercase, OR the next one is lowercase — the acronym has ended and this
//! capital begins the next word).
//!
//! camelCase shares none of that: lowercase each leading uppercase character
//! while `i == 0` or the NEXT character is also uppercase, then stop.
//! `HTTPResponse -> httpResponse` keeps the `R` for exactly that reason.

use std::sync::Arc;

use vybe_compiler::primitives::class_slots;
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::ops;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use super::object_fields::field_slot;

const TYPE_KEY: &str = "__type";

fn push_str(chunk: &mut Chunk, value: &str, line: u32) {
    chunk.emit_string_const(&Arc::from(value), line);
}

fn push_num(chunk: &mut Chunk, value: i32, line: u32) {
    chunk.emit_i32_const(value, line);
}

fn call(chunk: &mut Chunk, module: &str, func: &str, argc: u8, line: u32) {
    let idx = chunk.add_import(module, func);
    chunk.emit_call(idx, argc, line);
}

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn struct_set_drop(chunk: &mut Chunk, key: &str, line: u32) {
    class_slots::emit_class_set(
        chunk,
        class_slots::ObjSource::Stack,
        &field_slot(key),
        class_slots::ValueSource::Stack,
        line,
    );
}

/// `slot` holds an ASCII uppercase code unit — leaves an i32 condition.
fn push_is_upper(chunk: &mut Chunk, slot: u16, line: u32) {
    get(chunk, slot, line);
    push_num(chunk, 64, line);
    ops::emit_dyn_gt(chunk, line);
    get(chunk, slot, line);
    push_num(chunk, 91, line);
    ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
}

/// `slot` holds an ASCII lowercase code unit — leaves an i32 condition.
fn push_is_lower(chunk: &mut Chunk, slot: u16, line: u32) {
    get(chunk, slot, line);
    push_num(chunk, 96, line);
    ops::emit_dyn_gt(chunk, line);
    get(chunk, slot, line);
    push_num(chunk, 123, line);
    ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
}

/// `dst = index in [0,len) ? charCodeAt(text, index) : -1`.
///
/// ⛔ GUARDED ON PURPOSE. `wasm:js-string.charCodeAt` past the end THROWS here
/// (`index out of bounds`) instead of answering NaN the way JS does, so a bare
/// lookahead at the last character would abort the whole program. That cost a
/// program in the WebUtility work and is the same trap here.
fn read_char_guarded(chunk: &mut Chunk, text: u16, index: u16, len: u16, dst: u16, line: u32) {
    get(chunk, index, line);
    push_num(chunk, 0, line);
    ops::emit_dyn_ge(chunk, line);
    get(chunk, index, line);
    get(chunk, len, line);
    ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_value(line);
    get(chunk, text, line);
    get(chunk, index, line);
    call(chunk, "wasm:js-string", "charCodeAt", 2, line);
    chunk.emit_else(line);
    push_num(chunk, -1, line);
    chunk.emit_end(line);
    set(chunk, dst, line);
}

/// `out = concat(out, substring(text, i, i + 1))`.
fn append_char(chunk: &mut Chunk, out: u16, text: u16, i: u16, line: u32) {
    get(chunk, out, line);
    get(chunk, text, line);
    get(chunk, i, line);
    get(chunk, i, line);
    push_num(chunk, 1, line);
    ops::emit_dyn_add(chunk, line);
    call(chunk, "ecma:string", "substring", 3, line);
    call(chunk, "wasm:js-string", "concat", 2, line);
    set(chunk, out, line);
}

fn advance(chunk: &mut Chunk, i: u16, line: u32) {
    get(chunk, i, line);
    push_num(chunk, 1, line);
    ops::emit_dyn_add(chunk, line);
    set(chunk, i, line);
}

/// `ConvertName` for one separator policy. Receiver is local 0, the name local 1.
fn push_separator_convert_chunk(
    chunks: &mut Vec<Chunk>,
    sep: &str,
    upper: bool,
    line: u32,
) -> usize {
    let name = format!("__dotnet_json_convert_{}_{}", sep, upper);
    if let Some(idx) = chunks.iter().position(|chunk| chunk.name == name) {
        return idx;
    }
    let mut c = Chunk::new(&name);
    c.arity = 2;
    c.local_count = 2;
    let text = 1u16;
    let base = c.alloc_scratch(6);
    let (out, i, len, cur, prev, next) = (base, base + 1, base + 2, base + 3, base + 4, base + 5);

    get(&mut c, text, line);
    call(&mut c, "wasm:js-string", "length", 1, line);
    set(&mut c, len, line);
    push_str(&mut c, "", line);
    set(&mut c, out, line);
    push_num(&mut c, 0, line);
    set(&mut c, i, line);

    let guard = c.emit_block(line);
    let block = c.emit_block(line);
    let (loop_patch, _) = c.emit_loop_s(line);
    get(&mut c, i, line);
    get(&mut c, len, line);
    ops::emit_dyn_lt(&mut c, line);
    ops::emit_dyn_not(&mut c, line);
    ops::emit_dyn_to_bool(&mut c, line);
    c.emit_br_if(1, line);

    read_char_guarded(&mut c, text, i, len, cur, line);
    push_is_upper(&mut c, cur, line);
    c.emit_if(line);
    {
        // `i - 1` — no dyn subtract exists; negate the operand and add, which
        // is what `emit_dyn_neg` is there for.
        get(&mut c, i, line);
        push_num(&mut c, 1, line);
        ops::emit_dyn_neg(&mut c, line);
        ops::emit_dyn_add(&mut c, line);
        set(&mut c, prev, line);
        read_char_guarded(&mut c, text, prev, len, prev, line);
        get(&mut c, i, line);
        push_num(&mut c, 1, line);
        ops::emit_dyn_add(&mut c, line);
        set(&mut c, next, line);
        read_char_guarded(&mut c, text, next, len, next, line);

        get(&mut c, i, line);
        push_num(&mut c, 0, line);
        ops::emit_dyn_gt(&mut c, line);
        push_is_upper(&mut c, prev, line);
        c.emit_op(Op::I32_EQZ, line);
        push_is_lower(&mut c, next, line);
        c.emit_op(Op::I32_OR, line);
        c.emit_op(Op::I32_AND, line);
        c.emit_if(line);
        get(&mut c, out, line);
        push_str(&mut c, sep, line);
        call(&mut c, "wasm:js-string", "concat", 2, line);
        set(&mut c, out, line);
        c.emit_end(line);
    }
    c.emit_end(line);

    append_char(&mut c, out, text, i, line);
    advance(&mut c, i, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(loop_patch);
    c.emit_end(line);
    c.patch_block(block);
    c.emit_end(line);
    c.patch_block(guard);

    get(&mut c, out, line);
    call(
        &mut c,
        "ecma:string",
        if upper { "toUpperCase" } else { "toLowerCase" },
        1,
        line,
    );
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

/// `ConvertName` for camelCase.
fn push_camel_convert_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    const NAME: &str = "__dotnet_json_convert_camel";
    if let Some(idx) = chunks.iter().position(|chunk| chunk.name == NAME) {
        return idx;
    }
    let mut c = Chunk::new(NAME);
    c.arity = 2;
    c.local_count = 2;
    let text = 1u16;
    let base = c.alloc_scratch(6);
    let (out, i, len, cur, next, nx) = (base, base + 1, base + 2, base + 3, base + 4, base + 5);

    get(&mut c, text, line);
    call(&mut c, "wasm:js-string", "length", 1, line);
    set(&mut c, len, line);
    push_str(&mut c, "", line);
    set(&mut c, out, line);
    push_num(&mut c, 0, line);
    set(&mut c, i, line);

    let guard = c.emit_block(line);
    let block = c.emit_block(line);
    let (loop_patch, _) = c.emit_loop_s(line);
    get(&mut c, i, line);
    get(&mut c, len, line);
    ops::emit_dyn_lt(&mut c, line);
    ops::emit_dyn_not(&mut c, line);
    ops::emit_dyn_to_bool(&mut c, line);
    c.emit_br_if(1, line);

    read_char_guarded(&mut c, text, i, len, cur, line);
    get(&mut c, i, line);
    push_num(&mut c, 1, line);
    ops::emit_dyn_add(&mut c, line);
    set(&mut c, nx, line);
    read_char_guarded(&mut c, text, nx, len, next, line);

    // Stop as soon as the run ends.
    push_is_upper(&mut c, cur, line);
    get(&mut c, i, line);
    push_num(&mut c, 0, line);
    ops::emit_dyn_eq(&mut c, line);
    ops::emit_dyn_to_bool(&mut c, line);
    push_is_upper(&mut c, next, line);
    c.emit_op(Op::I32_OR, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_br_if(1, line);

    get(&mut c, out, line);
    get(&mut c, text, line);
    get(&mut c, i, line);
    get(&mut c, i, line);
    push_num(&mut c, 1, line);
    ops::emit_dyn_add(&mut c, line);
    call(&mut c, "ecma:string", "substring", 3, line);
    call(&mut c, "ecma:string", "toLowerCase", 1, line);
    call(&mut c, "wasm:js-string", "concat", 2, line);
    set(&mut c, out, line);
    advance(&mut c, i, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(loop_patch);
    c.emit_end(line);
    c.patch_block(block);
    c.emit_end(line);
    c.patch_block(guard);

    // Everything from the stopping point on is untouched.
    get(&mut c, out, line);
    get(&mut c, text, line);
    get(&mut c, i, line);
    get(&mut c, len, line);
    call(&mut c, "ecma:string", "substring", 3, line);
    call(&mut c, "wasm:js-string", "concat", 2, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

/// Which policy a static member names.
#[derive(Clone, Copy)]
pub enum Policy {
    SeparatorLower(&'static str),
    SeparatorUpper(&'static str),
    Camel,
}

/// `JsonNamingPolicy.<Name>` — a policy OBJECT carrying `ConvertName`.
///
/// A policy is a value in .NET (`JsonSerializerOptions.PropertyNamingPolicy`
/// holds one), so this mints an object rather than folding at the call site:
/// it has to survive being stored and read back before `ConvertName` is called.
pub fn emit_naming_policy(chunks: &mut Vec<Chunk>, current: usize, policy: Policy, line: u32) {
    let method_idx = match policy {
        Policy::SeparatorLower(sep) => push_separator_convert_chunk(chunks, sep, false, line),
        Policy::SeparatorUpper(sep) => push_separator_convert_chunk(chunks, sep, true, line),
        Policy::Camel => push_camel_convert_chunk(chunks, line),
    };
    let obj_slot = {
        let chunk = &mut chunks[current];
        let obj_slot = chunk.alloc_scratch(1);
        class_slots::emit_class_alloc(chunk, line);
        set(chunk, obj_slot, line);
        get(chunk, obj_slot, line);
        core_wasm::dup(chunk, line);
        chunk.emit_string_const(&Arc::from("JsonNamingPolicy"), line);
        struct_set_drop(chunk, TYPE_KEY, line);
        chunk.emit_op(Op::DROP, line);
        obj_slot
    };
    vybe_compiler::primitives::object::emit_bind_method(
        &mut chunks[current],
        obj_slot,
        "ConvertName",
        method_idx,
        line,
    );
    get(&mut chunks[current], obj_slot, line);
}

/// Drop a static member's arguments — these take none.
pub fn drop_args(chunk: &mut Chunk, argc: u8, line: u32) {
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
}

const _: Option<Value> = None;
