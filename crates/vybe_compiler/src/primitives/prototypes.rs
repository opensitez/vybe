//! JS prototype machinery — the ONE place for prototype-chain emission.
//!
//! ECMA-262 gives every function object a [[Prototype]] selected by its
//! kind (§20.2 %Function%, §27.7.1 %AsyncFunction%, §27.3.1
//! %GeneratorFunction%, §27.4.1 %AsyncGeneratorFunction%). The intrinsic
//! constructors are declared by the JS prelude (languages/js/mod.rs) as
//! runtime globals whose `.prototype` objects inherit from
//! %Function.prototype%, so `fn instanceof Function` holds through the
//! chain for every kind.
//!
//! Compiler call sites (function declarations in primitives/classes.rs,
//! lambdas in primitives/calls.rs, methods) must route through these
//! helpers rather than inlining raw opcodes.

use crate::primitives::class_slots;
use std::sync::Arc;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// The intrinsic constructor global whose `.prototype` becomes a function
/// object's [[Prototype]], selected by function kind.
pub fn fn_kind_intrinsic(is_async: bool, is_generator: bool) -> &'static str {
    match (is_async, is_generator) {
        (true, false) => "AsyncFunction",
        (false, true) => "GeneratorFunction",
        (true, true) => "AsyncGeneratorFunction",
        (false, false) => "Function",
    }
}

/// Stamp the function object on TOS with its kind's intrinsic prototype:
/// `fn.__proto__ = <Intrinsic>.prototype`.
///
/// Stack before: [fn]   Stack after: [] (consumed)
///
/// The intrinsic is read as a runtime global — the prelude declares all
/// four before any user function's metadata executes (top-level, hoisted).
pub fn emit_stamp_function_kind_proto(
    chunk: &mut Chunk,
    objects: bool,
    is_async: bool,
    is_generator: bool,
    line: u32,
) {
    // A language whose functions are not OBJECTS has no [[Prototype]] to
    // stamp. The operand is still consumed — the stack contract is the
    // helper's, not the object model's, and every caller has already
    // `dup`ed for it. See `Directives::functions_are_objects`.
    if !objects {
        chunk.emit_op(Op::DROP, line);
        return;
    }
    let intrinsic = fn_kind_intrinsic(is_async, is_generator);
    // §20.2.3.5 Function.prototype.toString reads this classifier to
    // synthesize the async/generator tokens (source text isn't retained).
    if is_async || is_generator {
        let kind = match (is_async, is_generator) {
            (true, false) => "async",
            (false, true) => "generator",
            _ => "async_generator",
        };
        crate::primitives::instructions::core_wasm::dup(chunk, line); // [fn, fn]
        chunk.emit_string_const(kind, line); // [fn, fn, kind]
        let slot = class_slots::resolve(
            &class_slots::ClassSlot::internal("__fn_kind"),
            &class_slots::PlainNames,
        );
        class_slots::emit_class_set(
            chunk,
            class_slots::ObjSource::Stack,
            &slot,
            class_slots::ValueSource::Stack,
            line,
        ); // [fn, fn]
    }
    crate::primitives::globals::emit_read(chunk, intrinsic, line); // [fn, ctor]
    let slot = class_slots::resolve(
        &class_slots::ClassSlot::Prototype,
        &class_slots::PlainNames,
    );
    class_slots::emit_class_get(
        chunk,
        class_slots::ObjSource::Stack,
        &slot,
        class_slots::Dest::Stack,
        line,
    ); // [fn, proto]
    let slot = class_slots::resolve(
        &class_slots::ClassSlot::ProtoLink,
        &class_slots::PlainNames,
    );
    class_slots::emit_class_set(
        chunk,
        class_slots::ObjSource::Stack,
        &slot,
        class_slots::ValueSource::Stack,
        line,
    ); // [fn]
}

/// §10.2.9 SetFunctionName / §10.2.10 SetFunctionLength: `name` and
/// `length` are non-enumerable data properties. Registers both in the
/// object's `__nonenum` set (the host convention `propertyIsEnumerable` /
/// `Object.keys` filter against).
///
/// Stack before: [fn]   Stack after: [] (consumed)
pub fn emit_stamp_fn_metadata_nonenum(chunk: &mut Chunk, objects: bool, line: u32) {
    // ⛔ Not a harmless extra for a language without function objects: this
    // emits a `struct.set` of `name` / `length` / `prototype` against a
    // declaration that has no such fields. In wast that alone stops the module
    // loading on a spec engine. See `Directives::functions_are_objects`.
    if !objects {
        chunk.emit_op(Op::DROP, line);
        return;
    }
    chunk.emit_string_const("name", line);
    chunk.emit_string_const("length", line);
    // §10.2.5: `prototype` on ordinary functions is non-enumerable too.
    chunk.emit_string_const("prototype", line);
    chunk.emit_array_new_fixed(0, 3, line); // [fn, [3 keys]]
    let slot = class_slots::resolve(
        &class_slots::ClassSlot::internal("__nonenum"),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_set(
        chunk,
        class_slots::ObjSource::Stack,
        &slot,
        class_slots::ValueSource::Stack,
        line,
    ); // [fn]
}
