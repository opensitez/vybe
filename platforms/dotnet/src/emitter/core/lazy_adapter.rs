//! `System.Lazy(Of T)` — deferred initialisation, shared by .NET languages.
//!
//! The type was absent from the .NET catalog entirely, so every `Lazy` member
//! answered "undefined is not callable" and the whole
//! `vb_lazy_thread_safe_mode_execution` category sat at 0/20.
//!
//! Shape: `{__type:"Lazy", __factory: <delegate>, __value: T, __created: Bool}`.
//! `Value` and `IsValueCreated` are READ-ONLY computed properties registered
//! through `tree_register::shared_emit_accessors`, which is what makes `.Value`
//! run the factory on first read instead of being an eager struct field —
//! `IsValueCreated` has to be `False` until someone asks.
//!
//! ## `isThreadSafe` is accepted and ignored, on purpose
//!
//! `New Lazy(Of T)(f, True)` and `New Lazy(Of T)(f, False)` differ only in
//! whether concurrent readers may race to run `f`. Both are required to publish
//! the SAME value, and the factory still runs exactly once per instance here.
//! The flag is dropped rather than stored because nothing can observe it.

use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use super::object_fields::field_slot;

const TYPE_KEY: &str = "__type";
const FACTORY_KEY: &str = "__factory";
const VALUE_KEY: &str = "__value";
const CREATED_KEY: &str = "__created";
const CREATING_KEY: &str = "__creating";

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

/// `New Lazy(Of T)(factory)` / `(factory, isThreadSafe)` / `()`.
///
/// Stack on entry: `[]`, `[factory]` or `[factory, isThreadSafe]` ;
/// on exit: `[lazy]`.
pub fn emit_lazy_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let (obj_slot, factory_slot) = {
        let chunk = &mut chunks[current];
        let base = chunk.alloc_scratch(2);
        (base, base + 1)
    };

    {
        let chunk = &mut chunks[current];
        if argc >= 2 {
            // `isThreadSafe` — see the module note.
            chunk.emit_op(Op::DROP, line);
        }
        if argc >= 1 {
            chunk.emit_op_u16(Op::LOCAL_SET, factory_slot, line);
        } else {
            // The parameterless overload default-constructs `T`, which needs
            // the TYPE ARGUMENT at run time. Nothing carries it here, so the
            // factory stays null and `Value` answers null rather than
            // inventing a value.
            chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
            chunk.emit_op_u16(Op::LOCAL_SET, factory_slot, line);
        }
    }
    if argc >= 1 {
        {
            let chunk = &mut chunks[current];
            lget(chunk, factory_slot, line);
            chunk.emit_op(Op::REF_IS_NULL, line);
            lget(chunk, factory_slot, line);
            let undef = chunk.add_import("wasm:js-undefined", "test");
            chunk.emit_call(undef, 1, line);
            chunk.emit_op(Op::I32_OR, line);
            chunk.emit_if(line);
        }
        crate::emitter::core::exceptions::emit_throw_typed(
            chunks,
            current,
            "ArgumentNullException",
            "valueFactory",
            line,
        );
        chunks[current].emit_end(line);
    }

    let obj_idx = chunks[current].add_import("ecma:object", "new");
    chunks[current].emit_call(obj_idx, 0, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);

        lget(chunk, obj_slot, line);
        chunk.emit_string_const("Lazy", line);
        class_slots::emit_class_set(
            chunk,
            ObjSource::Stack,
            &field_slot(TYPE_KEY),
            ValueSource::Stack,
            line,
        );

        lget(chunk, obj_slot, line);
        lget(chunk, factory_slot, line);
        class_slots::emit_class_set(
            chunk,
            ObjSource::Stack,
            &field_slot(FACTORY_KEY),
            ValueSource::Stack,
            line,
        );

        lget(chunk, obj_slot, line);
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        class_slots::emit_class_set(
            chunk,
            ObjSource::Stack,
            &field_slot(VALUE_KEY),
            ValueSource::Stack,
            line,
        );

        lget(chunk, obj_slot, line);
        chunk.emit_bool_const(false, line);
        class_slots::emit_class_set(
            chunk,
            ObjSource::Stack,
            &field_slot(CREATED_KEY),
            ValueSource::Stack,
            line,
        );

        lget(chunk, obj_slot, line);
        chunk.emit_bool_const(false, line);
        class_slots::emit_class_set(
            chunk,
            ObjSource::Stack,
            &field_slot(CREATING_KEY),
            ValueSource::Stack,
            line,
        );

        lget(chunk, obj_slot, line);
    }
}

/// `lazy.Value` — runs the factory on FIRST read and caches it.
///
/// Stack on entry: `[lazy]` ; on exit: `[value]`.
pub fn emit_lazy_value(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let obj_slot = {
        let chunk = &mut chunks[current];
        let slot = chunk.alloc_scratch(1);
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
        slot
    };

    {
        let chunk = &mut chunks[current];
        lget(chunk, obj_slot, line);
        class_slots::emit_class_get(
            chunk,
            ObjSource::Stack,
            &field_slot(CREATED_KEY),
            Dest::Stack,
            line,
        );
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_op(Op::I32_EQZ, line);
        // Void `if`: the arms only store to the object, nothing crosses the
        // block boundary.
        chunk.emit_if(line);

        lget(chunk, obj_slot, line);
        class_slots::emit_class_get(
            chunk,
            ObjSource::Stack,
            &field_slot(CREATING_KEY),
            Dest::Stack,
            line,
        );
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
    }
    crate::emitter::core::exceptions::emit_throw_typed(
        chunks,
        current,
        "InvalidOperationException",
        "ValueFactory attempted to access the Value property of this instance.",
        line,
    );
    {
        let chunk = &mut chunks[current];
        chunk.emit_end(line);

        lget(chunk, obj_slot, line);
        chunk.emit_bool_const(true, line);
        class_slots::emit_class_set(
            chunk,
            ObjSource::Stack,
            &field_slot(CREATING_KEY),
            ValueSource::Stack,
            line,
        );
    }

    let after_factory = chunks[current].emit_block(line);
    vybe_compiler::primitives::errors::emit_try_start(&mut chunks[current], line);

    {
        let chunk = &mut chunks[current];
        lget(chunk, obj_slot, line);
        lget(chunk, obj_slot, line);
        class_slots::emit_class_get(
            chunk,
            ObjSource::Stack,
            &field_slot(FACTORY_KEY),
            Dest::Stack,
            line,
        );
    }
    vybe_compiler::primitives::delegates::emit_invoke(chunks, current, 0, line);
    {
        let chunk = &mut chunks[current];
        class_slots::emit_class_set(
            chunk,
            ObjSource::Stack,
            &field_slot(VALUE_KEY),
            ValueSource::Stack,
            line,
        );

        lget(chunk, obj_slot, line);
        chunk.emit_bool_const(false, line);
        class_slots::emit_class_set(
            chunk,
            ObjSource::Stack,
            &field_slot(CREATING_KEY),
            ValueSource::Stack,
            line,
        );

        lget(chunk, obj_slot, line);
        chunk.emit_bool_const(true, line);
        class_slots::emit_class_set(
            chunk,
            ObjSource::Stack,
            &field_slot(CREATED_KEY),
            ValueSource::Stack,
            line,
        );
    }
    vybe_compiler::primitives::errors::emit_try_end(&mut chunks[current], line);
    chunks[current].emit_br(1, line);
    vybe_compiler::primitives::errors::emit_handler_block_end(&mut chunks[current], line);
    {
        let chunk = &mut chunks[current];
        let caught_slot = chunk.alloc_scratch(1);
        chunk.emit_op_u16(Op::LOCAL_SET, caught_slot, line);
        lget(chunk, obj_slot, line);
        chunk.emit_bool_const(false, line);
        class_slots::emit_class_set(
            chunk,
            ObjSource::Stack,
            &field_slot(CREATING_KEY),
            ValueSource::Stack,
            line,
        );
        chunk.emit_op_u16(Op::LOCAL_GET, caught_slot, line);
    }
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(after_factory);

    {
        let chunk = &mut chunks[current];
        chunk.emit_end(line);

        lget(chunk, obj_slot, line);
        class_slots::emit_class_get(
            chunk,
            ObjSource::Stack,
            &field_slot(VALUE_KEY),
            Dest::Stack,
            line,
        );
    }
}

/// `lazy.IsValueCreated`.
///
/// Stack on entry: `[lazy]` ; on exit: `[bool]`.
pub fn emit_lazy_is_value_created(chunks: &mut [Chunk], current: usize, line: u32) {
    class_slots::emit_class_get(
        &mut chunks[current],
        ObjSource::Stack,
        &field_slot(CREATED_KEY),
        Dest::Stack,
        line,
    );
}

/// `lazy.ToString()` — before the value is forced, .NET reports that the value
/// has not been created; afterwards it stringifies the cached value.
pub fn emit_lazy_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let obj_slot = {
        let chunk = &mut chunks[current];
        let slot = chunk.alloc_scratch(1);
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
        slot
    };

    {
        let chunk = &mut chunks[current];
        lget(chunk, obj_slot, line);
        class_slots::emit_class_get(
            chunk,
            ObjSource::Stack,
            &field_slot(CREATED_KEY),
            Dest::Stack,
            line,
        );
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if_value(line);

        lget(chunk, obj_slot, line);
        class_slots::emit_class_get(
            chunk,
            ObjSource::Stack,
            &field_slot(VALUE_KEY),
            Dest::Stack,
            line,
        );
        super::runtime_adapter::emit_tostring_dispatch(chunk, line);

        chunk.emit_else(line);
        chunk.emit_string_const("Value is not created.", line);
        chunk.emit_end(line);
    }
}
