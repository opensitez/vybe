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
//! Compiler call sites (function declarations in compiler/classes.rs,
//! lambdas in compiler/calls.rs, methods) must route through these
//! helpers rather than inlining raw opcodes.

use std::sync::Arc;
use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

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
    is_async: bool,
    is_generator: bool,
    line: u32,
) {
    let intrinsic = fn_kind_intrinsic(is_async, is_generator);
    // §20.2.3.5 Function.prototype.toString reads this classifier to
    // synthesize the async/generator tokens (source text isn't retained).
    if is_async || is_generator {
        let kind = match (is_async, is_generator) {
            (true, false) => "async",
            (false, true) => "generator",
            _ => "async_generator",
        };
        crate::instructions::core_wasm::dup(chunk, line); // [fn, fn]
        chunk.emit_string_const(kind, line); // [fn, fn, kind]
        let kind_key = chunk.add_constant(vybe_bytecode::Value::String(Arc::from("__fn_kind")));
        chunk.emit_op_u16(Op::STRUCT_SET, kind_key, line); // [fn, fn]
        chunk.emit_op(Op::DROP, line); // [fn]
    }
    let ctor_key = chunk.add_constant(vybe_bytecode::Value::String(Arc::from(intrinsic)));
    chunk.emit_op_u16(Op::GLOBAL_GET, ctor_key, line); // [fn, ctor]
    let proto_key = chunk.add_constant(vybe_bytecode::Value::String(Arc::from("prototype")));
    chunk.emit_op_u16(Op::STRUCT_GET, proto_key, line); // [fn, proto]
    let proto_link_key = chunk.add_constant(vybe_bytecode::Value::String(Arc::from("__proto__")));
    chunk.emit_op_u16(Op::STRUCT_SET, proto_link_key, line); // [fn]
    chunk.emit_op(Op::DROP, line);
}

/// §10.2.9 SetFunctionName / §10.2.10 SetFunctionLength: `name` and
/// `length` are non-enumerable data properties. Registers both in the
/// object's `__nonenum` set (the host convention `propertyIsEnumerable` /
/// `Object.keys` filter against).
///
/// Stack before: [fn]   Stack after: [] (consumed)
pub fn emit_stamp_fn_metadata_nonenum(chunk: &mut Chunk, line: u32) {
    chunk.emit_string_const("name", line);
    chunk.emit_string_const("length", line);
    // §10.2.5: `prototype` on ordinary functions is non-enumerable too.
    chunk.emit_string_const("prototype", line);
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 3, line); // [fn, [3 keys]]
    let key = chunk.add_constant(vybe_bytecode::Value::String(Arc::from("__nonenum")));
    chunk.emit_op_u16(Op::STRUCT_SET, key, line); // [fn]
    chunk.emit_op(Op::DROP, line);
}
