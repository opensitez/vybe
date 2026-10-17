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

const ARGUMENT_PLACE: &str = "__vybe_argument_place";
const ARGUMENT_VALUE: &str = "__vybe_argument_value";

/// Carry an evaluated value and its place until a late-bound signature is known.
/// The snapshot is separate from the place: a later argument can mutate the place.
pub fn emit_late_bound_argument(
    chunks: &mut [Chunk], current: usize, place: u16, value: u16, line: u32,
) {
    class_slots::emit_class_alloc(&mut chunks[current], line);
    for (name, slot) in [(ARGUMENT_PLACE, place), (ARGUMENT_VALUE, value)] {
        chunks[current].emit_dup(line);
        let key = class_slots::resolve_interned(&mut chunks[current],
            &class_slots::ClassSlot::internal(name), &class_slots::PlainNames);
        class_slots::emit_class_set(&mut chunks[current], class_slots::ObjSource::Stack,
            &key, class_slots::ValueSource::Local(slot), line);
    }
}

/// Select the signature's value or alias without changing ordinary arguments.
pub fn emit_bind_late_bound_argument(
    chunks: &mut [Chunk], current: usize, slot: u16, alias: bool, line: u32,
) {
    lget(&mut chunks[current], slot, line);
    recipes::is_object(&mut chunks[current], line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], slot, line);
    emit_field_get(chunks, current, ARGUMENT_PLACE, line);
    let place = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], place, line);
    emit_is_reference_from_slot(chunks, current, place, line);
    chunks[current].emit_if(line);
    if alias {
        lget(&mut chunks[current], place, line);
    } else {
        lget(&mut chunks[current], slot, line);
        emit_field_get(chunks, current, ARGUMENT_VALUE, line);
    }
    lset(&mut chunks[current], slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// A rest parameter contains argument carriers, not one carrier for the array.
pub fn emit_bind_late_bound_rest_arguments(
    chunks: &mut [Chunk], current: usize, slot: u16, alias: bool, line: u32,
) {
    let index = chunks[current].alloc_scratch(1);
    let length = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let len = chunks[current].add_import("ecma:array", "length");
    let get = chunks[current].add_import("ecma:array", "get");
    let set = chunks[current].add_import("ecma:array", "set");
    let to_i32 = chunks[current].add_import("wasm:js-number", "toI32");
    lget(&mut chunks[current], slot, line);
    chunks[current].emit_call(len, 1, line);
    chunks[current].emit_call(to_i32, 1, line);
    lset(&mut chunks[current], length, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], index, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    lget(&mut chunks[current], index, line);
    lget(&mut chunks[current], length, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    lget(&mut chunks[current], slot, line);
    lget(&mut chunks[current], index, line);
    chunks[current].emit_call(get, 2, line);
    lset(&mut chunks[current], value, line);
    emit_bind_late_bound_argument(chunks, current, value, alias, line);
    lget(&mut chunks[current], slot, line);
    lget(&mut chunks[current], index, line);
    lget(&mut chunks[current], value, line);
    chunks[current].emit_call(set, 3, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(&mut chunks[current], index, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
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

/// Annotate a reference without changing its storage identity or lookup key.
pub fn emit_byte_coordinate(chunks: &mut [Chunk], current: usize, offset: u64, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_dup(line);
    let key = class_slots::resolve_interned(
        chunk,
        &class_slots::ClassSlot::internal(super::pointers::REFERENCE_BYTE_OFFSET_KEY),
        &class_slots::PlainNames,
    );
    chunk.emit_f64_const(offset as f64, line);
    class_slots::emit_class_set(chunk, class_slots::ObjSource::Stack, &key,
        class_slots::ValueSource::Stack, line);
}

pub fn emit_byte_projection_metadata(
    chunks: &mut [Chunk], current: usize, root: Option<u16>, fields: u16, line: u32,
) {
    let chunk = &mut chunks[current];
    for (name, slot) in [
        (super::pointers::REFERENCE_BYTE_ROOT_KEY, root),
        (super::pointers::REFERENCE_BYTE_FIELDS_KEY, Some(fields)),
    ] {
        let Some(slot) = slot else { continue; };
        chunk.emit_dup(line);
        let key = class_slots::resolve_interned(chunk,
            &class_slots::ClassSlot::internal(name), &class_slots::PlainNames);
        lget(chunk, slot, line);
        class_slots::emit_class_set(chunk, class_slots::ObjSource::Stack,
            &key, class_slots::ValueSource::Stack, line);
    }
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
    emit_is_reference_from_slot(chunks, current, value, line);
}

/// Test an already-rooted value without extending the frame with another copy.
/// Keep the guarded tag read fresh: it may be backed by an observable getter.
fn emit_is_reference_from_slot(
    chunks: &mut [Chunk], current: usize, value: u16, line: u32,
) {
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
    chunk.emit_op_u16(Op::LOCAL_TEE, value, line);
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

/// Preserve the exact reference-read scratch ownership for a value whose
/// physical producer guarantees it is not an aggregate/reference wrapper.
/// This is not a dynamic test and must never be used for annotations or calls
/// to writable language helpers.
pub fn emit_nonreference_temporary_to_stack(chunks: &mut [Chunk], current: usize, line: u32) {
    // The inlined helper path reserves additional branch-local slots and
    // keeps a distinct frame-root region. Retain it exactly in that mode.
    if crate::primitives::polyfills::is_compiling_runtime_helper() {
        emit_autoderef_inline(chunks, current, line);
        return;
    }
    let chunk = &mut chunks[current];
    let value = chunk.autoderef_scratch();
    chunk.emit_op_u16(Op::LOCAL_TEE, value, line);
}

pub fn build_autoderef() -> Chunk {
    let mut chunk = Chunk::new("__stdlib_autoderef");
    chunk.arity = 1;
    chunk.local_count = 1;
    let mut chunks = vec![chunk];
    // The argument is already rooted in this helper's frame. Dispatch from
    // it directly instead of copying it into another frame-long local.
    emit_autoderef_from_slot(&mut chunks, 0, 0, 0);
    chunks[0].emit_op(Op::RETURN, 0);
    chunks.pop().unwrap()
}

/// Build the module-local read operation for a value already proven to be an
/// aggregate by its caller. This is a primitive body, not a polyfill or prelude.
/// The tag must still be read and type-tested here: the caller's earlier tag
/// read may have invoked a getter that returns a different value next time.
pub fn build_guarded_reference_read() -> Chunk {
    let mut chunk = Chunk::new("primitive.reference_read".to_string());
    chunk.arity = 1;
    chunk.local_count = 1;
    let mut chunks = vec![chunk];
    emit_guarded_reference_read_body(&mut chunks, 0, 0, 0);
    chunks[0].emit_op(Op::RETURN, 0);
    crate::primitives::bundle::mark_private_helper_body(&mut chunks[0]);
    chunks.pop().unwrap()
}

/// Emit the guarded operation from an existing local argument slot. The
/// direct-helper linker uses this form for its wide-function-index fallback.
pub fn emit_guarded_reference_read_body(
    chunks: &mut [Chunk], current: usize, object: u16, line: u32,
) {
    emit_autoderef_object_inline(chunks, current, object, line);
}

/// Inline fallback for the same operation when the module cannot address
/// another local function. Consumes one aggregate and leaves its read value.
/// No observable operation may intervene between the caller's aggregate proof
/// and this operation; the property reads themselves deliberately remain fresh.
pub fn emit_guarded_reference_read_inline(chunks: &mut [Chunk], current: usize, line: u32) {
    let first_temporary = chunks[current].local_count;
    let object = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], object, line);
    emit_guarded_reference_read_body(chunks, current, object, line);
    clear_guarded_read_temporaries(&mut chunks[current], first_temporary, line);
}

/// Wide-index fallback using the caller's existing aggregate slot. Its tag
/// and base temporaries have the same lifetime as a separate helper frame,
/// rather than retaining getter-produced objects until the caller returns.
pub fn emit_guarded_reference_read_inline_from_slot(
    chunks: &mut [Chunk], current: usize, object: u16, line: u32,
) {
    let first_temporary = chunks[current].local_count;
    emit_guarded_reference_read_body(chunks, current, object, line);
    clear_guarded_read_temporaries(&mut chunks[current], first_temporary, line);
}

fn clear_guarded_read_temporaries(chunk: &mut Chunk, first: u16, line: u32) {
    // The result stays on the operand stack while these stores discard only
    // private temporary roots. Never clear the caller's existing argument.
    for slot in first..chunk.local_count {
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        lset(chunk, slot, line);
    }
}

fn emit_autoderef_inline(chunks: &mut [Chunk], current: usize, line: u32) {
    let obj_slot = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], obj_slot, line);

    emit_autoderef_from_slot(chunks, current, obj_slot, line);
}

fn emit_autoderef_from_slot(
    chunks: &mut [Chunk], current: usize, obj_slot: u16, line: u32,
) {

    lget(&mut chunks[current], obj_slot, line);
    recipes::is_object(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    emit_autoderef_object_inline(chunks, current, obj_slot, line);

    chunks[current].emit_else(line);
    lget(&mut chunks[current], obj_slot, line);
    chunks[current].emit_end(line);
}

fn emit_autoderef_object_inline(
    chunks: &mut [Chunk],
    current: usize,
    obj_slot: u16,
    line: u32,
) {

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

    // A cell base supplies its value directly. All other bases share the
    // indexed read below. Keep the kind getter guarded by the object test.
    chunks[current].emit_block_typed(line, 1);
    lget(&mut chunks[current], base_slot, line);
    recipes::is_object(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_kind_eq(chunks, current, base_slot, CELL_KIND, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], base_slot, line);
    emit_cell_load(chunks, current, line);
    chunks[current].emit_br(2, line); // value exits both guards and this block
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

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
