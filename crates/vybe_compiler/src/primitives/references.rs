use crate::primitives::class_slots;
use std::sync::Arc;

use crate::primitives::pointers::{
    CARRAY_BASE_KEY, CARRAY_IDX_KEY, CARRAY_KIND, CELL_KIND, REF_KIND_KEY, REF_VALUE_KEY,
    SHARED_ADDR_KEY, SHARED_KIND,
};
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

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
    crate::primitives::bundle::emit_call_push_func(&mut chunks[current], "__vybe_futex_alloc16", line);
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
