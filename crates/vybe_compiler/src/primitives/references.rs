use crate::primitives::class_slots;

use crate::primitives::instructions::recipes;
use crate::primitives::pointers::{
    CARRAY_BASE_KEY, CARRAY_IDX_KEY, CARRAY_KIND, CELL_KIND, REF_KIND_KEY, REF_VALUE_KEY,
    SHARED_ADDR_KEY, SHARED_KIND,
};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn emit_field_get(chunks: &mut [Chunk], current: usize, field: &str, line: u32) {
    let key = class_slots::resolve(
        &class_slots::ClassSlot::internal(field),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_get(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &key,
        class_slots::Dest::Stack,
        line,
    );
}

fn emit_kind_eq(chunks: &mut [Chunk], current: usize, obj_slot: u16, kind_key: &str, line: u32) {
    lget(&mut chunks[current], obj_slot, line);
    emit_field_get(chunks, current, REF_KIND_KEY, line);
    let kind_slot = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], kind_slot, line);
    crate::primitives::ops::emit_string_slot_eq_literal(
        &mut chunks[current], kind_slot, kind_key, line,
    );
}

pub fn emit_cell_new(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) {
    class_slots::emit_class_alloc(&mut chunks[current], line);

    chunks[current].emit_dup(line);
    let kind_key = class_slots::resolve_interned(
        &mut chunks[current],
        &class_slots::ClassSlot::internal(REF_KIND_KEY),
        &class_slots::PlainNames,
    );
    chunks[current].emit_string_const(CELL_KIND, line);
    class_slots::emit_class_set(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &kind_key,
        class_slots::ValueSource::Stack,
        line,
    );

    chunks[current].emit_dup(line);
    let value_key = class_slots::resolve_interned(
        &mut chunks[current],
        &class_slots::ClassSlot::internal(REF_VALUE_KEY),
        &class_slots::PlainNames,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    class_slots::emit_class_set(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &value_key,
        class_slots::ValueSource::Stack,
        line,
    );
}

pub fn emit_cell_new_from_local(chunks: &mut [Chunk], current: usize, local_slot: u16, line: u32) {
    emit_cell_new(chunks, current, local_slot, line);
}

/// Consume a value and leave an i32 indicating whether it is a reference
/// wrapper. The decision must happen at runtime: a global can be promoted
/// after an earlier statement has already been compiled.
pub fn emit_is_reference_to_stack(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], value, line);
    lget(&mut chunks[current], value, line);
    recipes::is_object(&mut chunks[current], line);
    chunks[current].emit_if_i32(line);
    lget(&mut chunks[current], value, line);
    emit_field_get(chunks, current, REF_KIND_KEY, line);
    let test = chunks[current].add_import("wasm:js-string", "test");
    chunks[current].emit_call(test, 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_end(line);
}

/// Read through one PHP/Vybe reference wrapper if the value on the stack is
/// `{__ref_kind:"cell"}`, `{__ref_kind:"carray"}`, or `{__ref_kind:"shared"}`.
/// Non-reference values pass through unchanged.
///
/// This is the free-emitter sibling of the compiler's binding autoderef path.
/// Language adapters need it too: PHP's `global $x` binds `$x` to a carray
/// reference into the request-global object, and dynamic property helpers such
/// as `__php_field_get($x, "p")` must see the object stored in that slot, not
/// the reference wrapper itself.
pub fn emit_autoderef_to_stack(chunks: &mut [Chunk], current: usize, line: u32) {
    if crate::primitives::polyfills::is_compiling_runtime_helper() {
        emit_autoderef_inline(chunks, current, line);
        return;
    }
    let chunk = &mut chunks[current];
    let value = chunk.autoderef_scratch();
    lset(chunk, value, line);
    lget(chunk, value, line);
    recipes::is_object(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, value, line);
    emit_field_get(chunks, current, REF_KIND_KEY, line);
    let chunk = &mut chunks[current];
    let test = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(test, 1, line);
    chunk.emit_if_value(line);
    crate::primitives::globals::emit_read(chunk, "__vybe_autoderef", line);
    lget(chunk, value, line);
    crate::primitives::callable::emit_direct_invoke_chunk(chunk, 1, line);
    chunk.emit_else(line);
    lget(chunk, value, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    lget(chunk, value, line);
    chunk.emit_end(line);
}

pub fn build_autoderef() -> Chunk {
    let mut chunk = Chunk::new("__stdlib_autoderef");
    chunk.arity = 1;
    chunk.local_count = 1;
    lget(&mut chunk, 0, 0);
    let mut chunks = vec![chunk];
    emit_autoderef_inline(&mut chunks, 0, 0);
    chunks[0].emit_op(Op::RETURN, 0);
    chunks.pop().unwrap()
}

fn emit_autoderef_inline(chunks: &mut [Chunk], current: usize, line: u32) {
    let obj_slot = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], obj_slot, line);

    lget(&mut chunks[current], obj_slot, line);
    recipes::is_object(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    // Ordinary objects have no reference tag. Return them without walking
    // the cell/carray/shared dispatch ladder at every variable read.
    lget(&mut chunks[current], obj_slot, line);
    emit_field_get(chunks, current, REF_KIND_KEY, line);
    let kind_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_TEE, kind_slot, line);
    let test = chunks[current].add_import("wasm:js-string", "test");
    chunks[current].emit_call(test, 1, line);
    chunks[current].emit_if_value(line);

    // The wrapper's tag is stable throughout this read. Once its string type
    // is established, dispatch on that value without repeating field access
    // and type testing for each reference representation.
    let equals = chunks[current].add_import("wasm:js-string", "equals");

    lget(&mut chunks[current], kind_slot, line);
    chunks[current].emit_string_const(CELL_KIND, line);
    chunks[current].emit_call(equals, 2, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], obj_slot, line);
    emit_cell_load(chunks, current, line);
    chunks[current].emit_else(line);

    lget(&mut chunks[current], kind_slot, line);
    chunks[current].emit_string_const(CARRAY_KIND, line);
    chunks[current].emit_call(equals, 2, line);
    chunks[current].emit_if_value(line);

    let base_slot = chunks[current].alloc_scratch(1);
    lget(&mut chunks[current], obj_slot, line);
    emit_field_get(chunks, current, CARRAY_BASE_KEY, line);
    lset(&mut chunks[current], base_slot, line);

    lget(&mut chunks[current], base_slot, line);
    recipes::is_object(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    emit_kind_eq(chunks, current, base_slot, CELL_KIND, line);
    chunks[current].emit_if_params(line, 0, 1);
    lget(&mut chunks[current], base_slot, line);
    emit_cell_load(chunks, current, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], base_slot, line);
    lget(&mut chunks[current], obj_slot, line);
    emit_field_get(chunks, current, CARRAY_IDX_KEY, line);
    crate::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    lget(&mut chunks[current], base_slot, line);
    lget(&mut chunks[current], obj_slot, line);
    emit_field_get(chunks, current, CARRAY_IDX_KEY, line);
    crate::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);

    lget(&mut chunks[current], kind_slot, line);
    chunks[current].emit_string_const(SHARED_KIND, line);
    chunks[current].emit_call(equals, 2, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], obj_slot, line);
    emit_field_get(chunks, current, SHARED_ADDR_KEY, line);
    crate::primitives::threading::emit_atomic_load(&mut chunks[current], line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], obj_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    lget(&mut chunks[current], obj_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    lget(&mut chunks[current], obj_slot, line);
    chunks[current].emit_end(line);
}

/// `{__ref_kind:"shared", __addr}` — a word in SHARED linear memory, the
/// storage a WASM atomic acts on.
///
/// Allocates the word from the futex page (`__vybe_futex_alloc16`, which also
/// GROWS shared memory on first use — the `limit=0` half of the Interlocked
/// trap), `i32.atomic.store`s the value in `value_slot` as the word's initial
/// contents, and leaves the reference object on the stack. The binding then
/// holds this reference exactly as it would hold a cell: ordinary reads
/// autoderef through the `"shared"` arm (an atomic load), ordinary writes
/// store through it (an atomic store), and an atomic RMW asks for `__addr`.
pub fn emit_shared_word_new(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) {
    let addr_slot = chunks[current].alloc_scratch(1);
    crate::primitives::bundle::emit_call_push_func(
        &mut chunks[current],
        "__vybe_futex_alloc16",
        line,
    );
    crate::primitives::bundle::emit_call_invoke(&mut chunks[current], 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_TEE, addr_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    // The atomic store's i32 coercion handles f64-shaped values, the same
    // contract `emit_wasi_spawn` already relies on for the record words.
    crate::primitives::threading::emit_atomic_store(&mut chunks[current], line);

    class_slots::emit_class_alloc(&mut chunks[current], line);
    chunks[current].emit_dup(line);
    let kind_key = class_slots::resolve_interned(
        &mut chunks[current],
        &class_slots::ClassSlot::internal(REF_KIND_KEY),
        &class_slots::PlainNames,
    );
    chunks[current].emit_string_const(SHARED_KIND, line);
    class_slots::emit_class_set(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &kind_key,
        class_slots::ValueSource::Stack,
        line,
    );
    chunks[current].emit_dup(line);
    let addr_key = class_slots::resolve_interned(
        &mut chunks[current],
        &class_slots::ClassSlot::internal(SHARED_ADDR_KEY),
        &class_slots::PlainNames,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, addr_slot, line);
    class_slots::emit_class_set(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &addr_key,
        class_slots::ValueSource::Stack,
        line,
    );
}

/// A reference to a slot INSIDE a container: `{__ref_kind:"carray", __base, __idx}`.
///
/// `__idx` is a full `Value`, not an integer. The load and store both go through
/// the VM's polymorphic indexed access — `Op::ARRAY_GET` dispatches per
/// `ObjectKind` (Array by index, Map by Value key, plain Object by property
/// name) and `ecma:array.set` mirrors it — so `(base, key)` is already the one
/// representation a reference needs, whatever the container is. That is why
/// there is no third pointer kind here for member references.
///
/// The name is historical: the shape was introduced for c's decayed arrays, so
/// a numeric `__idx` was the only case. Nothing in the load/store path requires
/// it to be numeric.
pub fn emit_carray_new(
    chunks: &mut [Chunk],
    current: usize,
    base_slot: u16,
    key_slot: u16,
    line: u32,
) {
    class_slots::emit_class_alloc(&mut chunks[current], line);

    chunks[current].emit_dup(line);
    let kind_key = class_slots::resolve_interned(
        &mut chunks[current],
        &class_slots::ClassSlot::internal(REF_KIND_KEY),
        &class_slots::PlainNames,
    );
    chunks[current].emit_string_const(CARRAY_KIND, line);
    class_slots::emit_class_set(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &kind_key,
        class_slots::ValueSource::Stack,
        line,
    );

    chunks[current].emit_dup(line);
    let base_key = class_slots::resolve_interned(
        &mut chunks[current],
        &class_slots::ClassSlot::internal(CARRAY_BASE_KEY),
        &class_slots::PlainNames,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, base_slot, line);
    class_slots::emit_class_set(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &base_key,
        class_slots::ValueSource::Stack,
        line,
    );

    chunks[current].emit_dup(line);
    let idx_key = class_slots::resolve_interned(
        &mut chunks[current],
        &class_slots::ClassSlot::internal(CARRAY_IDX_KEY),
        &class_slots::PlainNames,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    class_slots::emit_class_set(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &idx_key,
        class_slots::ValueSource::Stack,
        line,
    );
}

pub fn emit_cell_load(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_key = class_slots::resolve(
        &class_slots::ClassSlot::internal(REF_VALUE_KEY),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_get(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &value_key,
        class_slots::Dest::Stack,
        line,
    );
}

pub fn emit_cell_store(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) {
    let value_key = class_slots::resolve_interned(
        &mut chunks[current],
        &class_slots::ClassSlot::internal(REF_VALUE_KEY),
        &class_slots::PlainNames,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    class_slots::emit_class_set(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &value_key,
        class_slots::ValueSource::Stack,
        line,
    );
}

/// Store through an existing reference, preserving the operand stack.
/// Compiler bindings and platform adapters share this dispatch: a container
/// reference's key may be a property name, not just an array index.
pub fn emit_store_through_pointer(
    chunks: &mut [Chunk],
    current: usize,
    pointer: u16,
    value: u16,
    line: u32,
) {
    lget(&mut chunks[current], pointer, line);
    recipes::is_object(&mut chunks[current], line);
    chunks[current].emit_if(line);

    emit_kind_eq(chunks, current, pointer, CELL_KIND, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], pointer, line);
    emit_cell_store(chunks, current, value, line);
    chunks[current].emit_else(line);

    emit_kind_eq(chunks, current, pointer, CARRAY_KIND, line);
    chunks[current].emit_if(line);
    let base = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    lget(&mut chunks[current], pointer, line);
    emit_field_get(chunks, current, CARRAY_BASE_KEY, line);
    lset(&mut chunks[current], base, line);
    lget(&mut chunks[current], pointer, line);
    emit_field_get(chunks, current, CARRAY_IDX_KEY, line);
    lset(&mut chunks[current], key, line);

    lget(&mut chunks[current], base, line);
    recipes::is_object(&mut chunks[current], line);
    chunks[current].emit_if_i32(line);
    emit_kind_eq(chunks, current, base, CELL_KIND, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], base, line);
    emit_cell_store(chunks, current, value, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], base, line);
    lget(&mut chunks[current], key, line);
    lget(&mut chunks[current], value, line);
    crate::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    emit_kind_eq(chunks, current, pointer, SHARED_KIND, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], pointer, line);
    emit_field_get(chunks, current, SHARED_ADDR_KEY, line);
    lget(&mut chunks[current], value, line);
    crate::primitives::threading::emit_atomic_store(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}
