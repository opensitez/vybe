use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, Dest, ObjSource, PlainNames, ValueSource,
};
use vybe_compiler::primitives::{globals, ops};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

const DELETED_NAME_GLOBAL: &str = "__py_deleted_name_singleton";
const DELETED_NAME_SLOT: &str = "__py_deleted_name";

pub fn emit_deleted_name(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let cached = chunk.alloc_scratch(1);

    globals::emit_read(chunk, DELETED_NAME_GLOBAL, line);
    chunk.emit_op_u16(Op::LOCAL_SET, cached, line);
    emit_slot_is_null_or_undefined(chunk, cached, line);
    chunk.emit_if_value(line);

    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, cached, line);

    chunk.emit_op_u16(Op::LOCAL_GET, cached, line);
    chunk.emit_string_const("DeletedName", line);
    let type_key = class_slots::resolve(&ClassSlot::TypeIdentity, &PlainNames);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &type_key, ValueSource::Stack, line);

    chunk.emit_op_u16(Op::LOCAL_GET, cached, line);
    chunk.emit_bool_const(true, line);
    let sentinel_key = class_slots::resolve(&ClassSlot::internal(DELETED_NAME_SLOT), &PlainNames);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &sentinel_key,
        ValueSource::Stack,
        line,
    );

    chunk.emit_op_u16(Op::LOCAL_GET, cached, line);
    globals::emit_write(chunk, DELETED_NAME_GLOBAL, line);
    chunk.emit_end(line);

    globals::emit_read(chunk, DELETED_NAME_GLOBAL, line);
}

pub fn emit_name_read(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);

    emit_is_deleted_name_slot(chunk, value, line);
    chunk.emit_if_value(line);
    crate::emitter::runtime_adapter::emit_throw_python_exception(
        chunk,
        "NameError",
        "name is not defined",
        line,
    );
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_end(line);
}

fn emit_slot_is_null_or_undefined(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    let is_undefined = chunk.add_import("wasm:js-undefined", "test");
    chunk.emit_call(is_undefined, 1, line);
    chunk.emit_op(Op::I32_OR, line);
}

fn emit_is_deleted_name_slot(chunk: &mut Chunk, slot: u16, line: u32) {
    let typeof_fn = chunk.add_import("ecma:value", "typeof");

    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_call(typeof_fn, 1, line);
    chunk.emit_string_const("object", line);
    ops::emit_dyn_eq(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_value(line);
    let key =
        class_slots::resolve_interned(chunk, &ClassSlot::internal(DELETED_NAME_SLOT), &PlainNames);
    class_slots::emit_class_get(chunk, ObjSource::Local(slot), &key, Dest::Stack, line);
    chunk.emit_bool_const(true, line);
    ops::emit_dyn_eq(chunk, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_end(line);
}
