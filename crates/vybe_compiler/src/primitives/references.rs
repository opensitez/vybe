use std::sync::Arc;

use crate::primitives::pointers::{
    CARRAY_BASE_KEY, CARRAY_IDX_KEY, CARRAY_KIND, CELL_KIND, REF_KIND_KEY, REF_VALUE_KEY,
};
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

pub fn emit_cell_new(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) {
    chunks[current].emit_struct_new(0, 0, line);

    chunks[current].emit_dup(line);
    let kind_key = chunks[current].add_constant(Value::String(Arc::from(REF_KIND_KEY)));
    chunks[current].emit_string_const(CELL_KIND, line);
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, kind_key, line);

    chunks[current].emit_dup(line);
    let value_key = chunks[current].add_constant(Value::String(Arc::from(REF_VALUE_KEY)));
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, value_key, line);
}

pub fn emit_cell_new_from_local(chunks: &mut [Chunk], current: usize, local_slot: u16, line: u32) {
    emit_cell_new(chunks, current, local_slot, line);
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
    chunks[current].emit_struct_new(0, 0, line);

    chunks[current].emit_dup(line);
    let kind_key = chunks[current].add_constant(Value::String(Arc::from(REF_KIND_KEY)));
    chunks[current].emit_string_const(CARRAY_KIND, line);
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, kind_key, line);

    chunks[current].emit_dup(line);
    let base_key = chunks[current].add_constant(Value::String(Arc::from(CARRAY_BASE_KEY)));
    chunks[current].emit_op_u16(Op::LOCAL_GET, base_slot, line);
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, base_key, line);

    chunks[current].emit_dup(line);
    let idx_key = chunks[current].add_constant(Value::String(Arc::from(CARRAY_IDX_KEY)));
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, idx_key, line);
}

pub fn emit_cell_load(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_key = chunks[current].add_constant(Value::String(Arc::from(REF_VALUE_KEY)));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, value_key, line);
}

pub fn emit_cell_store(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) {
    let value_key = chunks[current].add_constant(Value::String(Arc::from(REF_VALUE_KEY)));
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, value_key, line);
}
