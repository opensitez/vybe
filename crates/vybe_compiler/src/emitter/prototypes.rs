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

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;
use std::sync::Arc;

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
    let ctor_key = chunk.add_constant(vybe_bytecode::Value::String(Arc::from(intrinsic)));
    chunk.emit_op_u16(Op::GLOBAL_GET, ctor_key, line); // [fn, ctor]
    let proto_key = chunk.add_constant(vybe_bytecode::Value::String(Arc::from("prototype")));
    chunk.emit_op_u16(Op::STRUCT_GET, proto_key, line); // [fn, proto]
    let proto_link_key = chunk.add_constant(vybe_bytecode::Value::String(Arc::from("__proto__")));
    chunk.emit_op_u16(Op::STRUCT_SET, proto_link_key, line); // [fn]
    chunk.emit_op(Op::DROP, line);
}
