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
    chunks[current].emit_op(Op::ARRAY_GET, line);
    // Both getters are value reads. An alias stored in either namespace must
    // yield its referenced value; only the explicit *_ref APIs return places.
    vybe_compiler::primitives::references::emit_autoderef_to_stack(chunks, current, line);
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
    if module {
        lget(&mut chunks[current], store, line);
        lget(&mut chunks[current], key, line);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        let existing = chunks[current].alloc_scratch(1);
        lset(&mut chunks[current], existing, line);
        lget(&mut chunks[current], existing, line);
        vybe_compiler::primitives::references::emit_is_reference_to_stack(chunks, current, line);
        chunks[current].emit_if_value(line);
        lget(&mut chunks[current], existing, line);
        chunks[current].emit_else(line);
        vybe_compiler::primitives::references::emit_carray_new(chunks, current, store, key, line);
        chunks[current].emit_end(line);
    } else {
        vybe_compiler::primitives::references::emit_carray_new(chunks, current, store, key, line);
    }
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
    let store = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], store, line);
    lget(&mut chunks[current], store, line);
    lget(&mut chunks[current], key, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    let existing = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], existing, line);
    // Both branches capture the old referent before the write. Reuse the
    // existing reference test: plain values need no autoderef probe/helper.
    // This compiler spill is excluded by the existing program-root walk;
    // genuine source/global/reference owners remain roots.
    let old_value = chunks[current].alloc_scratch(1);
    lget(&mut chunks[current], existing, line);
    vybe_compiler::primitives::references::emit_is_reference_to_stack(chunks, current, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], existing, line);
    vybe_compiler::primitives::references::emit_autoderef_to_stack(chunks, current, line);
    lset(&mut chunks[current], old_value, line);
    vybe_compiler::primitives::references::emit_store_through_pointer(chunks, current, existing, value, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], existing, line);
    lset(&mut chunks[current], old_value, line);
    lget(&mut chunks[current], store, line);
    lget(&mut chunks[current], key, line);
    lget(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::ARRAY_SET, line);
    chunks[current].emit_end(line);

    // Namespace/reference writes must be visible before the reachability
    // check and before a destructor reenters PHP. Genuine aliases in globals,
    // include scopes and source locals remain roots; never special-case them
    // away. Keep this inline so no callee parameter becomes an extra owner.
    let chunk = &mut chunks[current];
    let destructor = vybe_compiler::primitives::class_slots::resolve_interned(
        chunk,
        &vybe_compiler::primitives::class_slots::ClassSlot::Slot(vybe_ast::ProtocolSlot::Destructor),
        &vybe_compiler::primitives::class_slots::PlainNames,
    );
    lget(chunk, old_value, line);
    lget(chunk, old_value, line);
    vybe_compiler::primitives::class_slots::emit_class_get(
        chunk,
        vybe_compiler::primitives::class_slots::ObjSource::Stack,
        &destructor,
        vybe_compiler::primitives::class_slots::Dest::Stack,
        line,
    );
    let finalize = chunk.add_import("php:object", "finalizeUnreachable");
    chunk.emit_call(finalize, 2, line);
    chunk.emit_op(Op::DROP, line);
    lget(&mut chunks[current], value, line);
}
