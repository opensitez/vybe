use crate::emitter::instructions::{core_wasm, host};
use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

use crate::emitter::{collections, dict, strings};

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[0].add_import(module, name);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(argc, line);
}

fn stash_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(argc as u16);
    for offset in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    base
}

pub fn emit_extend(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let src = base + 1;

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    core_wasm::dup(&mut chunks[current], line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    collections::emit_insert_range(chunks, current, line);
}

pub fn emit_get(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let recv = base;
    let key = base + 1;

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    dict::emit_get(chunks, current, line);

    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op(Op::DROP, line);
    if argc >= 3 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 2, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_end(line);
}

/// Python `gen.send(value)` — resume the generator with `value` (the result of
/// the pending `yield`) through the shared `generators.rs` `resume`, returning
/// the next yielded value. Same lazy layer JS/every language drives.
/// Stack: `[gen, value]` → `[yielded]`.
pub fn emit_gen_send(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let recv = base;
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
    crate::emitter::generators::emit_resume(&mut chunks[current], line);
}

/// Python `gen.throw(exc)` — resume the generator by throwing `exc` at the
/// pending `yield` via the shared `generators.rs` `resume_throw`.
/// Stack: `[gen, exc]` → `[yielded]`.
pub fn emit_gen_throw(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let recv = base;
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    } else {
        // throw() with no arg → GeneratorExit-ish; use a generic exception
        chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
        core_wasm::dup(&mut chunks[current], line);
        chunks[current].emit_string_const("", line);
        crate::emitter::errors::emit_exception_new_finalize(
            &mut chunks[current],
            "Exception",
            line,
        );
    }
    crate::emitter::generators::emit_resume_throw(&mut chunks[current], line);
}

/// Python `next(it[, default])`. For a generator, resume it through the shared
/// `generators.rs` machinery (`GEN_NEXT` → `[value, has_more]`) — the same lazy
/// path JS uses — so infinite generators advance one step instead of draining.
/// Non-generator iterables fall back to the shared `__vybe_pynext` helper.
pub fn emit_pynext(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let it = base;

    // if isGenerator(it)
    chunks[current].emit_op_u16(Op::LOCAL_GET, it, line);
    call_import(chunks, current, "ecma:value", "isGenerator", 1, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    // generator: GEN_NEXT → [value, has_more]
    chunks[current].emit_op_u16(Op::LOCAL_GET, it, line);
    crate::emitter::generators::emit_next(&mut chunks[current], line);
    let has_more = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    let value = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, has_more, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, has_more, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line); // has_more → value
    chunks[current].emit_else(line);
    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line); // exhausted → default
    } else {
        // exhausted, no default → raise StopIteration
        chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
        core_wasm::dup(&mut chunks[current], line);
        chunks[current].emit_string_const("", line);
        crate::emitter::errors::emit_exception_new_finalize(
            &mut chunks[current],
            "StopIteration",
            line,
        );
        crate::emitter::errors::emit_throw(&mut chunks[current], line);
        chunks[current].emit_op(Op::NULL, line); // unreachable (throw diverges)
    }
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    // not a generator → shared iterator-protocol next
    chunks[current].emit_op_u16(Op::LOCAL_GET, it, line);
    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    }
    collections::emit_runtime_helper_call(chunks, current, "__vybe_pynext", argc, line);
    chunks[current].emit_end(line);
}

/// Python from-end index normalization. Stack: `[obj, idx]` → `[normalized_idx]`.
///
/// `a[-1]` is "one from the end", not a real negative index: when `obj` is a
/// sequence (array or string) and `idx` is a negative number, this returns
/// `len(obj) + idx`; otherwise `idx` is returned unchanged so dict string/other
/// keys pass straight through. The `< 0` test is guarded behind an `isNumber`
/// check so a string key never hits numeric coercion — that guard is why this
/// replaces the shared `negative_index_wraps` flag (which trapped on `d['a']`).
pub fn emit_from_end(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let obj = base;
    let idx = base + 1;

    // if isNumber(idx)  (short-circuits the `< 0` test for string keys)
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    host::emit(&mut chunks[current], "wasm:js-number", "test", 1, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    // if idx < 0
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    crate::emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    // if isArray(obj) → len(obj) + idx
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_len_plus(chunks, current, obj, idx, line);
    chunks[current].emit_else(line);
    // else if isString(obj) → len(obj) + idx
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj, line);
    host::emit(&mut chunks[current], "wasm:js-string", "test", 1, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_len_plus(chunks, current, obj, idx, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line); // idx >= 0 → unchanged
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line); // not a number → unchanged (dict/other key)
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_end(line);
}

/// Python `x in y` membership. Stack: `[container, needle]` → `[bool]`.
/// string → substring test; array → element test; else (dict/object) → own key.
/// Set literals are lowered to `.has()` upstream and never reach here.
pub fn emit_contains(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let test_str = chunk.add_import("wasm:js-string", "test");
    let str_includes = chunk.add_import("ecma:string", "includes");
    let is_array = chunk.add_import("ecma:array", "isArray");
    let arr_includes = chunk.add_import("ecma:array", "includes");
    let has_own = chunk.add_import("ecma:object", "hasOwn");

    let needle = chunk.alloc_scratch(1);
    let container = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, needle, line);
    chunk.emit_op_u16(Op::LOCAL_SET, container, line);

    // string → substring test
    chunk.emit_op_u16(Op::LOCAL_GET, container, line);
    chunk.emit_call(test_str, 1, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, container, line);
    chunk.emit_op_u16(Op::LOCAL_GET, needle, line);
    chunk.emit_call(str_includes, 2, line);
    chunk.emit_else(line);

    // array → element test
    chunk.emit_op_u16(Op::LOCAL_GET, container, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, container, line);
    chunk.emit_op_u16(Op::LOCAL_GET, needle, line);
    chunk.emit_call(arr_includes, 2, line);
    chunk.emit_else(line);

    // object → user `__contains__(self, item)` if present, else own-key test
    let contains_key = chunk
        .add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("__contains__")));
    let contains_method = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, container, line);
    chunk.emit_op_u16(Op::STRUCT_GET, contains_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, contains_method, line);
    chunk.emit_op_u16(Op::LOCAL_GET, contains_method, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line); // 1 if a __contains__ method is present
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, contains_method, line);
    chunk.emit_op_u16(Op::LOCAL_GET, container, line);
    chunk.emit_op_u16(Op::LOCAL_GET, needle, line);
    chunk.emit_op_u8(Op::CALL_REF, 2, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, container, line);
    chunk.emit_op_u16(Op::LOCAL_GET, needle, line);
    chunk.emit_call(has_own, 2, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// `len(obj) + idx` — helper for the from-end wrap. Stack: `[]` → `[value]`.
fn emit_len_plus(chunks: &mut [Chunk], current: usize, obj: u16, idx: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj, line);
    emit_length(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    crate::emitter::ops::emit_dyn_add(&mut chunks[current], line);
}

pub fn emit_index(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let recv = base;
    let needle = base + 1;

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    host::emit(&mut chunks[current], "wasm:js-string", "test", 1, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, needle, line);
    strings::emit_index_of(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, needle, line);
    collections::emit_index_of(chunks, current, line);
    chunks[current].emit_end(line);
}

pub fn emit_clear(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let recv = base;
    let keys_key =
        chunks[current].add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("__keys")));

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, keys_key, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    collections::emit_len(chunks, current, line);
    collections::emit_remove_range(chunks, current, line);

    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:set", "clear", 1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    dict::emit_method_clear_stack(chunks, current, line);
    chunks[current].emit_end(line);
}

pub fn emit_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    call_import(chunks, current, "ecma:set", "add", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

fn emit_remove_impl(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let value = base + 1;

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    crate::emitter::ops::emit_dyn_ne(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    call_import(chunks, current, "ecma:array", "removeValue", 2, line);

    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    call_import(chunks, current, "ecma:set", "delete", 2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_remove(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_remove_impl(chunks, current, line);
}

pub fn emit_discard(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_remove_impl(chunks, current, line);
}

pub fn emit_copy(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let recv = base;
    let keys_key =
        chunks[current].add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("__keys")));

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, keys_key, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    crate::emitter::ops::emit_dyn_ne(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    collections::emit_len(chunks, current, line);
    collections::emit_slice(chunks, current, line);

    chunks[current].emit_else(line);
    call_import(chunks, current, "ecma:set", "new", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:set", "union", 2, line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    dict::emit_new(chunks, current, line);
    let out = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:object", "assign", 2, line);
    chunks[current].emit_end(line);
}

pub fn emit_update(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let src = base + 1;
    let keys_key =
        chunks[current].add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("__keys")));

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, keys_key, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    call_import(chunks, current, "ecma:set", "unionWith", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    call_import(chunks, current, "ecma:object", "assign", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_end(line);
    chunks[current].emit_op(Op::NULL, line);
}

fn emit_set_update_call(chunks: &mut [Chunk], current: usize, func: &str, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    call_import(chunks, current, "ecma:set", func, 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_intersection_update(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_set_update_call(chunks, current, "intersectWith", line);
}

pub fn emit_difference_update(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_set_update_call(chunks, current, "exceptWith", line);
}

pub fn emit_symmetric_difference_update(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_set_update_call(chunks, current, "symmetricExceptWith", line);
}

pub fn emit_pop(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let recv = base;
    let value_slot = chunks[current].local_count;
    chunks[current].alloc_scratch(1);

    if argc == 1 {
        let index_slot = chunks[current].local_count;
        chunks[current].alloc_scratch(1);

        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        collections::emit_len(chunks, current, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_SUB, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
        collections::emit_remove_at(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        return;
    } else {
        // Dispatch on isArray, NOT `__keys` presence: a Python dict is a plain
        // JS object with no `__keys`, so a keys-based check misclassifies it as
        // a list. `list.pop(i)` splices; `dict.pop(k[, default])` reads the
        // value then removes the property natively via `ecma:object.delete`.
        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        call_import(chunks, current, "ecma:array", "isArray", 1, line);
        crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);

        // list.pop(i): value = recv[i]; remove_at(recv, i); value
        let index = base + 1;
        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
        collections::emit_remove_at(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);

        chunks[current].emit_else(line);

        // dict.pop(k[, default]): value = recv[k]
        let key = base + 1;
        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
        dict::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if(line);
        // missing key → default (or null)
        if argc >= 3 {
            chunks[current].emit_op_u16(Op::LOCAL_GET, base + 2, line);
        } else {
            chunks[current].emit_op(Op::NULL, line);
        }
        chunks[current].emit_else(line);
        // present → delete the property natively and return the value
        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
        call_import(chunks, current, "ecma:object", "delete", 2, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunks[current].emit_end(line);

        chunks[current].emit_end(line);
    }
}

pub fn emit_length(chunks: &mut [Chunk], current: usize, line: u32) {
    // Polymorphic `len`: string → char length, array → element count,
    // Set/Map → `.size`, otherwise (dict/object) → `Object.keys(o).length`.
    // Uses the object's native property enumeration (a Python dict IS a JS
    // object) — no `__keys` array, so literal and built dicts count the same.
    let base = stash_args(chunks, current, 1, line);
    let recv = base;
    let size_key =
        chunks[current].add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("size")));

    // User-defined `__len__` → call it with the receiver. (Cross-language:
    // bound alongside `__get_length`/`__get_count`.)
    let len_key = chunks[current]
        .add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("__len__")));
    let len_method = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, len_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_method, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_method, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line); // 1 if a method is present
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_method, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_else(line);

    // isString(recv) → string length
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    host::emit(&mut chunks[current], "wasm:js-string", "test", 1, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_else(line);

    // isArray(recv) → element count
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_else(line);

    // isView(recv) → typed-array (bytes) length
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:arraybuffer", "isView", 1, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:uint8array", "length", 1, line);
    chunks[current].emit_else(line);

    // has `.size` (Set/Map) → use it, else Object.keys(recv).length
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, size_key, line);
    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op(Op::DROP, line); // drop null `.size`
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line); // close the isView (bytes) branch
    chunks[current].emit_end(line); // close the __len__ dispatch branch
}
