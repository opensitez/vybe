//! Runtime extraction into an explicit PHP symbol table.
//!
//! Scope entries are generic reference cells supplied by the walker. No VM
//! frame inspection is needed. Arguments: array, flags, prefix, scope.

use vybe_compiler::primitives::{loops, ops, references};
use vybe_runtime::{opcode::Op, Chunk};

fn get(c: &mut Chunk, slot: u16, line: u32) { c.emit_op_u16(Op::LOCAL_GET, slot, line); }
fn set(c: &mut Chunk, slot: u16, line: u32) { c.emit_op_u16(Op::LOCAL_SET, slot, line); }
fn number(c: &mut Chunk, n: f64, line: u32) { c.emit_f64_const(n, line); }
fn call(c: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let import = c.add_import(module, name);
    c.emit_call(import, argc, line);
}
fn equal(c: &mut Chunk, slot: u16, n: f64, line: u32) {
    get(c, slot, line);
    number(c, n, line);
    ops::emit_dyn_eq(c, line);
    ops::emit_dyn_to_bool(c, line);
}

pub fn emit_scope_new(chunks: &mut [Chunk], current: usize, line: u32) {
    call(&mut chunks[current], "ecma:map", "new", 0, line);
}

pub fn emit_scope_cell(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    references::emit_autoderef_to_stack(chunks, current, line);
    super::copy_adapter::emit_php_copy_on_assign(chunks, current, 1, line);
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    references::emit_cell_new(chunks, current, value, line);
}

pub fn emit_scope_ref(chunks: &mut [Chunk], current: usize, line: u32) {
    let c = &mut chunks[current];
    let key = c.alloc_scratch(1);
    let scope = c.alloc_scratch(1);
    let value = c.alloc_scratch(1);
    let pointer = c.alloc_scratch(1);
    set(c, key, line);
    set(c, scope, line);
    get(c, scope, line);
    get(c, key, line);
    call(c, "ecma:map", "has", 2, line);
    ops::emit_dyn_to_bool(c, line);
    c.emit_if_value(line);
    get(c, scope, line);
    get(c, key, line);
    c.emit_op(Op::ARRAY_GET, line);
    set(c, value, line);
    get(c, value, line);
    references::emit_is_reference_to_stack(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_else(line);
    references::emit_cell_new(chunks, current, value, line);
    let c = &mut chunks[current];
    set(c, pointer, line);
    get(c, scope, line);
    get(c, key, line);
    get(c, pointer, line);
    c.emit_op(Op::ARRAY_SET, line);
    get(c, pointer, line);
    c.emit_end(line);
    c.emit_else(line);
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    set(c, value, line);
    references::emit_cell_new(chunks, current, value, line);
    let c = &mut chunks[current];
    set(c, pointer, line);
    get(c, scope, line);
    get(c, key, line);
    get(c, pointer, line);
    c.emit_op(Op::ARRAY_SET, line);
    get(c, pointer, line);
    c.emit_end(line);
}
fn between(c: &mut Chunk, slot: u16, lo: f64, hi: f64, line: u32) {
    get(c, slot, line);
    number(c, lo, line);
    ops::emit_dyn_ge(c, line);
    ops::emit_dyn_to_bool(c, line);
    get(c, slot, line);
    number(c, hi, line);
    ops::emit_dyn_le(c, line);
    ops::emit_dyn_to_bool(c, line);
    c.emit_op(Op::I32_AND, line);
}

/// Leave an i32 indicating whether a string is a PHP variable identifier.
fn valid_name(chunks: &mut [Chunk], current: usize, name: u16, line: u32) {
    let c = &mut chunks[current];
    let i = c.alloc_scratch(1);
    let len = c.alloc_scratch(1);
    let ch = c.alloc_scratch(1);
    let valid = c.alloc_scratch(1);
    get(c, name, line);
    call(c, "ecma:string", "length", 1, line);
    set(c, len, line);
    get(c, len, line);
    number(c, 0.0, line);
    ops::emit_dyn_gt(c, line);
    ops::emit_dyn_to_bool(c, line);
    set(c, valid, line);
    number(c, 0.0, line);
    set(c, i, line);
    let state = loops::emit_loop_start(chunks, current, line);
    let c = &mut chunks[current];
    get(c, i, line);
    get(c, len, line);
    ops::emit_dyn_lt(c, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);
    let c = &mut chunks[current];
    get(c, name, line);
    get(c, i, line);
    call(c, "ecma:string", "charCodeAt", 2, line);
    set(c, ch, line);
    between(c, ch, 65.0, 90.0, line);
    between(c, ch, 97.0, 122.0, line);
    c.emit_op(Op::I32_OR, line);
    equal(c, ch, 95.0, line);
    c.emit_op(Op::I32_OR, line);
    get(c, ch, line);
    number(c, 128.0, line);
    ops::emit_dyn_ge(c, line);
    ops::emit_dyn_to_bool(c, line);
    c.emit_op(Op::I32_OR, line);
    between(c, ch, 48.0, 57.0, line);
    get(c, i, line);
    number(c, 0.0, line);
    ops::emit_dyn_gt(c, line);
    ops::emit_dyn_to_bool(c, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_op(Op::I32_OR, line);
    get(c, valid, line);
    c.emit_op(Op::I32_AND, line);
    set(c, valid, line);
    get(c, i, line);
    number(c, 1.0, line);
    ops::emit_dyn_add(c, line);
    set(c, i, line);
    loops::emit_loop_end(chunks, current, state, line);
    get(&mut chunks[current], valid, line);
}

pub fn emit_extract_scope(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let c = &mut chunks[current];
    let scope = c.alloc_scratch(1);
    let prefix = c.alloc_scratch(1);
    let flags = c.alloc_scratch(1);
    let array = c.alloc_scratch(1);
    let mode = c.alloc_scratch(1);
    let refs = c.alloc_scratch(1);
    let keys = c.alloc_scratch(1);
    let len = c.alloc_scratch(1);
    let i = c.alloc_scratch(1);
    let key = c.alloc_scratch(1);
    let target = c.alloc_scratch(1);
    let exists = c.alloc_scratch(1);
    let valid = c.alloc_scratch(1);
    let accept = c.alloc_scratch(1);
    let prefixed = c.alloc_scratch(1);
    let value = c.alloc_scratch(1);
    let pointer = c.alloc_scratch(1);
    let count = c.alloc_scratch(1);
    set(c, scope, line);
    set(c, prefix, line);
    set(c, flags, line);
    set(c, array, line);
    get(c, flags, line);
    number(c, 256.0, line);
    ops::emit_dyn_ge(c, line);
    ops::emit_dyn_to_bool(c, line);
    set(c, refs, line);
    get(c, refs, line);
    c.emit_if_value(line);
    get(c, flags, line);
    number(c, -256.0, line);
    ops::emit_dyn_add(c, line);
    c.emit_else(line);
    get(c, flags, line);
    c.emit_end(line);
    set(c, mode, line);
    between(c, mode, 0.0, 6.0, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    super::type_guard::emit_throw_const(chunks, current, "ValueError",
        "extract(): Argument #2 ($flags) must be a valid extract type", line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], array, line);
    references::emit_autoderef_to_stack(chunks, current, line);
    set(&mut chunks[current], array, line);
    super::array_adapter::emit_php_key_list_from_slot(chunks, current, array, line);
    let c = &mut chunks[current];
    set(c, keys, line);
    get(c, keys, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    set(c, len, line);
    number(c, 0.0, line);
    set(c, i, line);
    number(c, 0.0, line);
    set(c, count, line);
    let state = loops::emit_loop_start(chunks, current, line);
    let c = &mut chunks[current];
    get(c, i, line);
    get(c, len, line);
    ops::emit_dyn_lt(c, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);
    let c = &mut chunks[current];
    get(c, keys, line);
    get(c, i, line);
    c.emit_op(Op::ARRAY_GET, line);
    set(c, key, line);
    get(c, key, line);
    vybe_compiler::primitives::convert::emit_to_string(c, line);
    set(c, target, line);
    get(c, scope, line);
    get(c, target, line);
    call(c, "ecma:map", "has", 2, line);
    ops::emit_dyn_to_bool(c, line);
    set(c, exists, line);
    valid_name(chunks, current, target, line);
    let c = &mut chunks[current];
    set(c, valid, line);
    c.emit_i32_const(1, line);
    set(c, accept, line);
    c.emit_i32_const(0, line);
    set(c, prefixed, line);
    equal(c, mode, 1.0, line);
    c.emit_if(line);
    get(c, exists, line);
    c.emit_op(Op::I32_EQZ, line);
    set(c, accept, line);
    c.emit_end(line);
    equal(c, mode, 2.0, line);
    c.emit_if(line);
    get(c, exists, line);
    set(c, prefixed, line);
    c.emit_end(line);
    equal(c, mode, 3.0, line);
    c.emit_if(line);
    c.emit_i32_const(1, line);
    set(c, prefixed, line);
    c.emit_end(line);
    equal(c, mode, 4.0, line);
    c.emit_if(line);
    get(c, valid, line);
    c.emit_op(Op::I32_EQZ, line);
    set(c, prefixed, line);
    c.emit_end(line);
    equal(c, mode, 5.0, line);
    equal(c, mode, 6.0, line);
    c.emit_op(Op::I32_OR, line);
    c.emit_if(line);
    get(c, exists, line);
    set(c, accept, line);
    equal(c, mode, 5.0, line);
    set(c, prefixed, line);
    c.emit_end(line);
    get(c, prefixed, line);
    c.emit_if(line);
    get(c, prefix, line);
    c.emit_string_const("_", line);
    call(c, "ecma:string", "concat", 2, line);
    get(c, target, line);
    call(c, "ecma:string", "concat", 2, line);
    set(c, target, line);
    c.emit_end(line);
    valid_name(chunks, current, target, line);
    let c = &mut chunks[current];
    get(c, accept, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_if(line);
    get(c, target, line);
    c.emit_string_const("this", line);
    ops::emit_dyn_eq(c, line);
    ops::emit_dyn_to_bool(c, line);
    c.emit_if(line);
    super::type_guard::emit_throw_const(chunks, current, "Error", "Cannot re-assign $this", line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], refs, line);
    chunks[current].emit_if(line);
    references::emit_carray_new(chunks, current, array, key, line);
    set(&mut chunks[current], value, line);
    chunks[current].emit_else(line);
    let c = &mut chunks[current];
    get(c, array, line);
    get(c, key, line);
    c.emit_op(Op::ARRAY_GET, line);
    references::emit_autoderef_to_stack(chunks, current, line);
    super::copy_adapter::emit_php_copy_on_assign(chunks, current, 1, line);
    set(&mut chunks[current], value, line);
    chunks[current].emit_end(line);
    let c = &mut chunks[current];
    get(c, scope, line);
    get(c, target, line);
    call(c, "ecma:map", "has", 2, line);
    ops::emit_dyn_to_bool(c, line);
    // EXTR_REFS rebinds the local symbol rather than writing through a
    // pre-existing alias (for example a by-reference function parameter).
    get(c, refs, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_if(line);
    get(c, scope, line);
    get(c, target, line);
    c.emit_op(Op::ARRAY_GET, line);
    set(c, pointer, line);
    references::emit_store_through_pointer(chunks, current, pointer, value, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], refs, line);
    chunks[current].emit_if_value(line);
    // Keep the array-element reference itself as the symbol's binding.
    // Wrapping it in another cell would make assignment replace the alias.
    get(&mut chunks[current], value, line);
    chunks[current].emit_else(line);
    references::emit_cell_new(chunks, current, value, line);
    chunks[current].emit_end(line);
    let c = &mut chunks[current];
    set(c, pointer, line);
    get(c, scope, line);
    get(c, target, line);
    get(c, pointer, line);
    c.emit_op(Op::ARRAY_SET, line);
    c.emit_end(line);
    get(c, count, line);
    number(c, 1.0, line);
    ops::emit_dyn_add(c, line);
    set(c, count, line);
    c.emit_end(line);
    get(c, i, line);
    number(c, 1.0, line);
    ops::emit_dyn_add(c, line);
    set(c, i, line);
    loops::emit_loop_end(chunks, current, state, line);
    get(&mut chunks[current], count, line);
}
