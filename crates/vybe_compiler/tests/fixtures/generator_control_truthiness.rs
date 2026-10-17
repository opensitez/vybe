// Frozen original entry/resume control emission.
use vybe_compiler::primitives::instructions::{host, recipes};
use vybe_compiler::primitives::{class_slots, ops};
use vybe_runtime::{Chunk, Op};

pub fn emit_entry_control(
    chunk: &mut Chunk,
    control_slot: u16,
    line: u32,
    emit_return: &mut dyn FnMut(&mut Chunk, u32),
) {
    chunk.emit_op_u16(Op::LOCAL_GET, control_slot, line);
    recipes::is_object(chunk, line);
    // `ref.test` already yields an i32; `Op::IF` consumes it directly.
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, control_slot, line);
    let slot = class_slots::resolve(
        &class_slots::ClassSlot::internal("__vybe_generator_control"),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_get(
        chunk,
        class_slots::ObjSource::Stack,
        &slot,
        class_slots::Dest::Stack,
        line,
    );
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, control_slot, line);
    let op_key = class_slots::resolve_interned(
        chunk,
        &class_slots::ClassSlot::internal("op"),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_get(
        chunk,
        class_slots::ObjSource::Stack,
        &op_key,
        class_slots::Dest::Stack,
        line,
    );
    chunk.emit_string_const("throw", line);
    host::emit(chunk, "wasm:js-string", "equals", 2, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, control_slot, line);
    let value_key = class_slots::resolve_interned(
        chunk,
        &class_slots::ClassSlot::internal("value"),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_get(
        chunk,
        class_slots::ObjSource::Stack,
        &value_key,
        class_slots::Dest::Stack,
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(chunk, line);

    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, control_slot, line);
    class_slots::emit_class_get(
        chunk,
        class_slots::ObjSource::Stack,
        &op_key,
        class_slots::Dest::Stack,
        line,
    );
    chunk.emit_string_const("return", line);
    host::emit(chunk, "wasm:js-string", "equals", 2, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, control_slot, line);
    class_slots::emit_class_get(
        chunk,
        class_slots::ObjSource::Stack,
        &value_key,
        class_slots::Dest::Stack,
        line,
    );
    emit_return(chunk, line);

    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

pub fn emit_resume_dispatch(
    chunk: &mut Chunk,
    line: u32,
    emit_return: &mut dyn FnMut(&mut Chunk, u32),
) -> (u16, u16) {
    let resume_slot = chunk.local_count;
    chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, resume_slot, line);

    let result_slot = chunk.local_count;
    chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, resume_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, resume_slot, line);
    recipes::is_object(chunk, line);
    // `ref.test` already yields an i32; `Op::IF` consumes it directly.
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, resume_slot, line);
    let slot = class_slots::resolve(
        &class_slots::ClassSlot::internal("__vybe_generator_control"),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_get(
        chunk,
        class_slots::ObjSource::Stack,
        &slot,
        class_slots::Dest::Stack,
        line,
    );
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, resume_slot, line);
    let op_key = class_slots::resolve_interned(
        chunk,
        &class_slots::ClassSlot::internal("op"),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_get(
        chunk,
        class_slots::ObjSource::Stack,
        &op_key,
        class_slots::Dest::Stack,
        line,
    );
    chunk.emit_string_const("throw", line);
    host::emit(chunk, "wasm:js-string", "equals", 2, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, resume_slot, line);
    let value_key = class_slots::resolve_interned(
        chunk,
        &class_slots::ClassSlot::internal("value"),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_get(
        chunk,
        class_slots::ObjSource::Stack,
        &value_key,
        class_slots::Dest::Stack,
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(chunk, line);

    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, resume_slot, line);
    class_slots::emit_class_get(
        chunk,
        class_slots::ObjSource::Stack,
        &op_key,
        class_slots::Dest::Stack,
        line,
    );
    chunk.emit_string_const("return", line);
    host::emit(chunk, "wasm:js-string", "equals", 2, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, resume_slot, line);
    class_slots::emit_class_get(
        chunk,
        class_slots::ObjSource::Stack,
        &value_key,
        class_slots::Dest::Stack,
        line,
    );
    emit_return(chunk, line);

    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);

    (resume_slot, result_slot)
}
