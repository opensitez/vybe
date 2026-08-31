//! `System.Text.Json.JsonDocument` / `JsonElement`.
//!
//! A `JsonElement` is a WRAPPER around a parsed value, not the value itself:
//! `.ValueKind` has to answer without a call, and `.GetProperty(...)` has to
//! return something that answers `.GetInt32()` in turn. So `wrap` stamps the
//! kind as a FIELD and binds the accessors as methods; every accessor that
//! yields a child re-wraps, which is what makes chaining work.
//!
//! ⛔ `ValueKind` cannot ride the tree. `PropertyDef::getter` only accepts a
//! `HostTarget`, so there is no Common-emitted instance property, and C# does
//! not set `member_invokes_parameterless_method` (only Pascal does) — so a
//! zero-arity method would never be invoked by a bare read either.
//!
//! Kinds measured against .NET SDK 10 — note `true` and `false` are SEPARATE
//! kinds, so the shared six-tag model has to be split on the way out:
//!
//! ```text
//!   {…} -> Object    […] -> Array    "s" -> String
//!   7   -> Number    true -> True    false -> False    null -> Null
//! ```

use std::sync::Arc;

use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::{json, ops};
use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

use super::object_fields::field_slot;

const VALUE_KEY: &str = "__json";

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn push_str(chunk: &mut Chunk, value: &str, line: u32) {
    chunk.emit_string_const(&Arc::from(value), line);
}

fn call(chunk: &mut Chunk, module: &str, func: &str, argc: u8, line: u32) {
    let idx = chunk.add_import(module, func);
    chunk.emit_call(idx, argc, line);
}

fn field_get(chunk: &mut Chunk, key: &str, line: u32) {
    class_slots::emit_class_get(chunk, ObjSource::Stack, &field_slot(key), Dest::Stack, line);
}

fn field_set_drop(chunk: &mut Chunk, key: &str, line: u32) {
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(key),
        ValueSource::Stack,
        line,
    );
}

/// Push the receiver's wrapped value. Receiver is local 0 in every bound method.
fn push_inner(chunk: &mut Chunk, line: u32) {
    get(chunk, 0, line);
    field_get(chunk, VALUE_KEY, line);
}

/// Call the `wrap` chunk on the value currently on the stack.
fn call_wrap(chunk: &mut Chunk, wrap_idx: usize, line: u32) {
    let arg = chunk.alloc_scratch(1);
    set(chunk, arg, line);
    chunk.emit_op_u16(Op::REF_FUNC, wrap_idx as u16, line);
    chunk.emit(0, line);
    get(chunk, arg, line);
    chunk.emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
}

/// `wrap(value) -> JsonElement`.
fn push_wrap_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    const NAME: &str = "__dotnet_json_element_wrap";
    if let Some(idx) = chunks.iter().position(|c| c.name == NAME) {
        return idx;
    }
    // Reserve the slot so the accessors can reference `wrap` recursively.
    chunks.push(Chunk::new(NAME));
    let wrap_idx = chunks.len() - 1;

    let accessors: Vec<(&str, usize)> = vec![
        ("GetProperty", push_get_property_chunk(chunks, wrap_idx, line)),
        ("TryGetProperty", push_try_get_property_chunk(chunks, wrap_idx, line)),
        ("GetInt32", push_identity_chunk(chunks, "__dotnet_json_get_int32", line)),
        ("GetInt64", push_identity_chunk(chunks, "__dotnet_json_get_int32", line)),
        ("GetDouble", push_identity_chunk(chunks, "__dotnet_json_get_int32", line)),
        ("GetDecimal", push_identity_chunk(chunks, "__dotnet_json_get_int32", line)),
        ("GetString", push_identity_chunk(chunks, "__dotnet_json_get_int32", line)),
        ("GetBoolean", push_identity_chunk(chunks, "__dotnet_json_get_int32", line)),
        ("GetArrayLength", push_array_length_chunk(chunks, line)),
        ("EnumerateArray", push_enumerate_array_chunk(chunks, wrap_idx, line)),
        ("EnumerateObject", push_enumerate_object_chunk(chunks, line)),
        ("GetRawText", push_raw_text_chunk(chunks, line)),
        ("Clone", push_clone_chunk(chunks, wrap_idx, line)),
    ];
    let to_string = push_to_string_chunk(chunks, line);

    let mut c = Chunk::new(NAME);
    c.arity = 1;
    c.local_count = 1;
    let value = 0u16;
    let base = c.alloc_scratch(2);
    let (obj, kind) = (base, base + 1);

    json::emit_value_kind(&mut c, value, line);
    set(&mut c, kind, line);

    class_slots::emit_class_alloc(&mut c, line);
    set(&mut c, obj, line);
    get(&mut c, obj, line);
    core_wasm::dup(&mut c, line);
    push_str(&mut c, "JsonElement", line);
    field_set_drop(&mut c, "__type", line);
    core_wasm::dup(&mut c, line);
    get(&mut c, value, line);
    field_set_drop(&mut c, VALUE_KEY, line);

    // The six shared tags, split into .NET's eight names.
    for spelling in ["ValueKind", "valuekind"] {
        core_wasm::dup(&mut c, line);
        get(&mut c, kind, line);
        push_str(&mut c, "boolean", line);
        ops::emit_dyn_eq(&mut c, line);
        ops::emit_dyn_to_bool(&mut c, line);
        c.emit_if_value(line);
        get(&mut c, value, line);
        ops::emit_dyn_to_bool(&mut c, line);
        c.emit_if_value(line);
        push_str(&mut c, "True", line);
        c.emit_else(line);
        push_str(&mut c, "False", line);
        c.emit_end(line);
        c.emit_else(line);
        // Capitalise the tag: object -> Object, and so on.
        for (tag, dotnet) in [
            ("object", "Object"),
            ("array", "Array"),
            ("string", "String"),
            ("number", "Number"),
        ] {
            get(&mut c, kind, line);
            push_str(&mut c, tag, line);
            ops::emit_dyn_eq(&mut c, line);
            ops::emit_dyn_to_bool(&mut c, line);
            c.emit_if_value(line);
            push_str(&mut c, dotnet, line);
            c.emit_else(line);
        }
        push_str(&mut c, "Null", line);
        for _ in 0..4 {
            c.emit_end(line);
        }
        c.emit_end(line);
        field_set_drop(&mut c, spelling, line);
    }
    c.emit_op(Op::DROP, line);

    // ⛔ THE INDEXER CANNOT BE A BOUND METHOD. `element[1]` compiles to a plain
    // key read on the WRAPPER, which would look for a field named `1` — so an
    // array wrapper stamps each already-wrapped child under its own index. The
    // children are wrapped eagerly, which is the price of `[i]` answering
    // something that still responds to `.GetInt32()`.
    {
        let idx_base = c.alloc_scratch(3);
        let (i, n, child) = (idx_base, idx_base + 1, idx_base + 2);
        get(&mut c, kind, line);
        push_str(&mut c, "array", line);
        ops::emit_dyn_eq(&mut c, line);
        ops::emit_dyn_to_bool(&mut c, line);
        c.emit_if(line);
        get(&mut c, value, line);
        call(&mut c, "ecma:array", "length", 1, line);
        set(&mut c, n, line);
        c.emit_i32_const(0, line);
        set(&mut c, i, line);
        let guard = c.emit_block(line);
        let block = c.emit_block(line);
        let (loop_patch, _) = c.emit_loop_s(line);
        get(&mut c, i, line);
        get(&mut c, n, line);
        ops::emit_dyn_lt(&mut c, line);
        ops::emit_dyn_not(&mut c, line);
        ops::emit_dyn_to_bool(&mut c, line);
        c.emit_br_if(1, line);
        get(&mut c, value, line);
        get(&mut c, i, line);
        c.emit_op(Op::ARRAY_GET, line);
        call_wrap(&mut c, wrap_idx, line);
        set(&mut c, child, line);
        get(&mut c, obj, line);
        get(&mut c, i, line);
        get(&mut c, child, line);
        c.emit_op(Op::ARRAY_SET, line);
        get(&mut c, i, line);
        c.emit_i32_const(1, line);
        ops::emit_dyn_add(&mut c, line);
        set(&mut c, i, line);
        c.emit_br(0, line);
        c.emit_end(line);
        c.patch_loop(loop_patch);
        c.emit_end(line);
        c.patch_block(block);
        c.emit_end(line);
        c.patch_block(guard);
        c.emit_end(line);
    }

    for (name, idx) in accessors {
        vybe_compiler::primitives::object::emit_bind_method(&mut c, obj, name, idx, line);
    }
    vybe_compiler::primitives::object::emit_bind_method(
        &mut c,
        obj,
        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::ToString),
        to_string,
        line,
    );
    get(&mut c, obj, line);
    c.emit_op(Op::RETURN, line);
    chunks[wrap_idx] = c;
    wrap_idx
}

/// Every scalar getter is the same read — .NET distinguishes them by the type
/// it returns, and this runtime carries one numeric type, so `GetInt32` and
/// `GetDouble` genuinely are the same operation here.
fn push_identity_chunk(chunks: &mut Vec<Chunk>, name: &str, line: u32) -> usize {
    if let Some(idx) = chunks.iter().position(|c| c.name == name) {
        return idx;
    }
    let mut c = Chunk::new(name);
    c.arity = 1;
    c.local_count = 1;
    push_inner(&mut c, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

fn push_get_property_chunk(chunks: &mut Vec<Chunk>, wrap_idx: usize, line: u32) -> usize {
    const NAME: &str = "__dotnet_json_get_property";
    if let Some(idx) = chunks.iter().position(|c| c.name == NAME) {
        return idx;
    }
    let mut c = Chunk::new(NAME);
    c.arity = 2;
    c.local_count = 2;
    push_inner(&mut c, line);
    get(&mut c, 1, line);
    c.emit_op(Op::ARRAY_GET, line);
    call_wrap(&mut c, wrap_idx, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

fn push_try_get_property_chunk(chunks: &mut Vec<Chunk>, _wrap: usize, line: u32) -> usize {
    const NAME: &str = "__dotnet_json_try_get_property";
    if let Some(idx) = chunks.iter().position(|c| c.name == NAME) {
        return idx;
    }
    let mut c = Chunk::new(NAME);
    c.arity = 2;
    c.local_count = 2;
    push_inner(&mut c, line);
    get(&mut c, 1, line);
    c.emit_op(Op::ARRAY_GET, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_op(Op::I32_EQZ, line);
    ops::emit_i32_to_bool(&mut c, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

fn push_array_length_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    const NAME: &str = "__dotnet_json_array_length";
    if let Some(idx) = chunks.iter().position(|c| c.name == NAME) {
        return idx;
    }
    let mut c = Chunk::new(NAME);
    c.arity = 1;
    c.local_count = 1;
    push_inner(&mut c, line);
    call(&mut c, "ecma:array", "length", 1, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

/// `EnumerateArray()` — the wrapped elements, so `e.GetInt32()` works in the loop.
fn push_enumerate_array_chunk(chunks: &mut Vec<Chunk>, wrap_idx: usize, line: u32) -> usize {
    const NAME: &str = "__dotnet_json_enumerate_array";
    if let Some(idx) = chunks.iter().position(|c| c.name == NAME) {
        return idx;
    }
    let mut c = Chunk::new(NAME);
    c.arity = 1;
    c.local_count = 1;
    let base = c.alloc_scratch(4);
    let (src, out, i, n) = (base, base + 1, base + 2, base + 3);
    push_inner(&mut c, line);
    set(&mut c, src, line);
    call(&mut c, "ecma:array", "new", 0, line);
    set(&mut c, out, line);
    get(&mut c, src, line);
    call(&mut c, "ecma:array", "length", 1, line);
    set(&mut c, n, line);
    c.emit_i32_const(0, line);
    set(&mut c, i, line);

    let guard = c.emit_block(line);
    let block = c.emit_block(line);
    let (loop_patch, _) = c.emit_loop_s(line);
    get(&mut c, i, line);
    get(&mut c, n, line);
    ops::emit_dyn_lt(&mut c, line);
    ops::emit_dyn_not(&mut c, line);
    ops::emit_dyn_to_bool(&mut c, line);
    c.emit_br_if(1, line);
    get(&mut c, out, line);
    get(&mut c, src, line);
    get(&mut c, i, line);
    c.emit_op(Op::ARRAY_GET, line);
    call_wrap(&mut c, wrap_idx, line);
    call(&mut c, "ecma:array", "push", 2, line);
    c.emit_op(Op::DROP, line);
    get(&mut c, i, line);
    c.emit_i32_const(1, line);
    ops::emit_dyn_add(&mut c, line);
    set(&mut c, i, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(loop_patch);
    c.emit_end(line);
    c.patch_block(block);
    c.emit_end(line);
    c.patch_block(guard);
    get(&mut c, out, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

/// `EnumerateObject()` — objects carrying `Name`, which is what the corpus reads.
fn push_enumerate_object_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    const NAME: &str = "__dotnet_json_enumerate_object";
    if let Some(idx) = chunks.iter().position(|c| c.name == NAME) {
        return idx;
    }
    let mut c = Chunk::new(NAME);
    c.arity = 1;
    c.local_count = 1;
    let base = c.alloc_scratch(5);
    let (src, keys, out, i, n) = (base, base + 1, base + 2, base + 3, base + 4);
    push_inner(&mut c, line);
    set(&mut c, src, line);
    get(&mut c, src, line);
    call(&mut c, "ecma:object", "keys", 1, line);
    set(&mut c, keys, line);
    call(&mut c, "ecma:array", "new", 0, line);
    set(&mut c, out, line);
    get(&mut c, keys, line);
    call(&mut c, "ecma:array", "length", 1, line);
    set(&mut c, n, line);
    c.emit_i32_const(0, line);
    set(&mut c, i, line);

    let guard = c.emit_block(line);
    let block = c.emit_block(line);
    let (loop_patch, _) = c.emit_loop_s(line);
    get(&mut c, i, line);
    get(&mut c, n, line);
    ops::emit_dyn_lt(&mut c, line);
    ops::emit_dyn_not(&mut c, line);
    ops::emit_dyn_to_bool(&mut c, line);
    c.emit_br_if(1, line);
    get(&mut c, out, line);
    class_slots::emit_class_alloc(&mut c, line);
    core_wasm::dup(&mut c, line);
    get(&mut c, keys, line);
    get(&mut c, i, line);
    c.emit_op(Op::ARRAY_GET, line);
    field_set_drop(&mut c, "Name", line);
    core_wasm::dup(&mut c, line);
    get(&mut c, keys, line);
    get(&mut c, i, line);
    c.emit_op(Op::ARRAY_GET, line);
    field_set_drop(&mut c, "name", line);
    call(&mut c, "ecma:array", "push", 2, line);
    c.emit_op(Op::DROP, line);
    get(&mut c, i, line);
    c.emit_i32_const(1, line);
    ops::emit_dyn_add(&mut c, line);
    set(&mut c, i, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(loop_patch);
    c.emit_end(line);
    c.patch_block(block);
    c.emit_end(line);
    c.patch_block(guard);
    get(&mut c, out, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

fn push_raw_text_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    const NAME: &str = "__dotnet_json_raw_text";
    if let Some(idx) = chunks.iter().position(|c| c.name == NAME) {
        return idx;
    }
    let mut c = Chunk::new(NAME);
    c.arity = 1;
    c.local_count = 1;
    push_inner(&mut c, line);
    call(&mut c, "ecma:json", "stringify", 1, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

/// ⛔ `ToString()` on a STRING element is the string itself, not `"…"` — only a
/// non-string renders as JSON text.
fn push_to_string_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    const NAME: &str = "__dotnet_json_element_tostring";
    if let Some(idx) = chunks.iter().position(|c| c.name == NAME) {
        return idx;
    }
    let mut c = Chunk::new(NAME);
    c.arity = 1;
    c.local_count = 1;
    let v = c.alloc_scratch(1);
    push_inner(&mut c, line);
    set(&mut c, v, line);
    get(&mut c, v, line);
    let idx = c.add_import("wasm:js-string", "test");
    c.emit_call(idx, 1, line);
    c.emit_if_value(line);
    get(&mut c, v, line);
    c.emit_else(line);
    get(&mut c, v, line);
    call(&mut c, "ecma:json", "stringify", 1, line);
    c.emit_end(line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

fn push_clone_chunk(chunks: &mut Vec<Chunk>, _wrap: usize, line: u32) -> usize {
    const NAME: &str = "__dotnet_json_clone";
    if let Some(idx) = chunks.iter().position(|c| c.name == NAME) {
        return idx;
    }
    let mut c = Chunk::new(NAME);
    c.arity = 1;
    c.local_count = 1;
    get(&mut c, 0, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

/// `JsonDocument.Parse(text)` — a document whose `RootElement` is the wrapped
/// tree. `Dispose` exists because the corpus writes `using var doc = …`.
pub fn emit_document_parse(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let wrap_idx = push_wrap_chunk(chunks, line);
    let dispose = push_clone_chunk(chunks, wrap_idx, line);
    let chunk = &mut chunks[current];
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let base = chunk.alloc_scratch(2);
    let (root, obj) = (base, base + 1);
    call(chunk, "ecma:json", "parse", 1, line);
    call_wrap(chunk, wrap_idx, line);
    set(chunk, root, line);

    class_slots::emit_class_alloc(chunk, line);
    set(chunk, obj, line);
    get(chunk, obj, line);
    core_wasm::dup(chunk, line);
    push_str(chunk, "JsonDocument", line);
    field_set_drop(chunk, "__type", line);
    for spelling in ["RootElement", "rootelement"] {
        core_wasm::dup(chunk, line);
        get(chunk, root, line);
        field_set_drop(chunk, spelling, line);
    }
    chunk.emit_op(Op::DROP, line);
    vybe_compiler::primitives::object::emit_bind_method(chunk, obj, "Dispose", dispose, line);
    get(chunk, obj, line);
}
