//! PHP request/global-variable storage.
//!
//! Runtime PHP includes compile as separate modules, so the shared AST
//! `GlobalNamespace` primitive only sees the current module's declared names.
//! PHP's `$GLOBALS` / `global $x` semantics are request-wide. Keep that
//! PHP-specific spelling here and back it with the ECMA `globalThis` singleton
//! so dynamically included modules share storage without teaching the shared
//! compiler about PHP.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}

fn emit_is_null_or_undefined(chunk: &mut Chunk, slot: u16, line: u32) {
    lget(chunk, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    lget(chunk, slot, line);
    let undef = chunk.add_import("wasm:js-undefined", "test");
    chunk.emit_call(undef, 1, line);
    chunk.emit_op(Op::I32_OR, line);
}

/// Leaves the shared PHP globals object on the stack.
pub fn emit_php_globals_object(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "ecma:globalThis", "get", 0, line);
    call_import(chunks, current, "php:globals", "store", 1, line);
}

/// `__php_global_get($name)` -> value-or-null.
pub fn emit_php_global_get(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_get(chunks, current, argc, line, false);
}

pub fn emit_php_module_get(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_get(chunks, current, argc, line, true);
}

fn emit_variable_store(chunks: &mut [Chunk], current: usize, line: u32, module: bool) {
    if !module {
        emit_php_globals_object(chunks, current, line);
        return;
    }
    call_import(chunks, current, "vybe:php", "include_scope", 0, line);
    let scope = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], scope, line);
    emit_is_null_or_undefined(&mut chunks[current], scope, line);
    chunks[current].emit_if_value(line);
    emit_php_globals_object(chunks, current, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], scope, line);
    chunks[current].emit_end(line);
}

fn emit_get(chunks: &mut [Chunk], current: usize, argc: u8, line: u32, module: bool) {
    if argc == 0 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let key = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], key, line);

    emit_variable_store(chunks, current, line, module);
    lget(&mut chunks[current], key, line);
    call_import(chunks, current, "ecma:object", "get", 2, line);
    let value = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], value, line);

    emit_is_null_or_undefined(&mut chunks[current], value, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], value, line);
    chunks[current].emit_end(line);
}

/// `__php_global_ref($name)` -> reference to the request-global slot.
pub fn emit_php_global_ref(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_ref(chunks, current, argc, line, false);
}

pub fn emit_php_module_ref(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_ref(chunks, current, argc, line, true);
}

fn emit_ref(chunks: &mut [Chunk], current: usize, argc: u8, line: u32, module: bool) {
    if argc == 0 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let key = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], key, line);

    emit_variable_store(chunks, current, line, module);
    let store = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], store, line);

    vybe_compiler::primitives::references::emit_carray_new(chunks, current, store, key, line);
}

/// `__php_global_set($name, $value)` -> `$value`.
pub fn emit_php_global_set(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_set(chunks, current, argc, line, false);
}

pub fn emit_php_module_set(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_set(chunks, current, argc, line, true);
}

fn emit_set(chunks: &mut [Chunk], current: usize, argc: u8, line: u32, module: bool) {
    if argc < 2 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let value = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], value, line);
    lset(&mut chunks[current], key, line);

    emit_variable_store(chunks, current, line, module);
    lget(&mut chunks[current], key, line);
    lget(&mut chunks[current], value, line);
    call_import(chunks, current, "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(&mut chunks[current], value, line);
}
