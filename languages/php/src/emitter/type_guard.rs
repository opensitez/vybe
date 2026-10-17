//! Runtime argument type-guards for PHP builtins.
//!
//! PHP errors are constructed by the same PHP core class callables as source
//! `new TypeError(...)`, not by allocating a bare object and stamping a name.
//! Their inherited methods, default fields and Throwable ancestry therefore
//! come from ordinary native class lowering. Constructor dependencies must be
//! selected by the PHP bootstrap before these emitted cold paths can execute.
//!
//! Guards spill arguments once and restore them unchanged on the valid path.

use vybe_compiler::primitives::class_slots::{self};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// What an argument is required to be for a builtin to accept it.
#[derive(Clone, Copy)]
pub enum Expect {
    /// Must be a PHP array; a non-array (scalar or object) is illegal.
    /// (`array_merge`, `sort`, `array_keys`, `in_array` haystack, …)
    Array,
    /// Must be coercible to string/number; an array is illegal but scalars are
    /// fine (`strlen`, `str_starts_with`, `abs`, `round`).
    NotArray,
    /// Scalars or objects implementing PHP's string-conversion protocol.
    StringCoercible,
    /// Must be a collection or object; a bare scalar is illegal but objects
    /// (Countable, Traversable, generators) are fine (`count`,
    /// `iterator_to_array`).
    NotScalar,
}

/// Construct the real PHP core exception and throw it. There is deliberately
/// no undefined-constructor/bare-object fallback: source bootstrap dependency
/// selection must provide the same class binding as an explicit `new`.
pub fn emit_throw_const(
    chunks: &mut [Chunk],
    current: usize,
    exc_name: &str,
    msg: &str,
    line: u32,
) {
    chunks[current].emit_string_const(msg, line);
    emit_throw_from_stack(chunks, current, exc_name, line);
}

/// Consume one message. Preserve the previous finalizer's exactly-once String
/// conversion, then call the actual three-argument PHP class constructor. Its
/// normal class lowering owns methods, field defaults and Error/Exception's
/// distinct ancestry; reflection stamps are not a substitute for that class.
pub fn emit_throw_from_stack(
    chunks: &mut [Chunk],
    current: usize,
    exc_name: &str,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let message_slot = chunk.alloc_scratch(1);
    let stringify = chunk.add_import("ecma:string", "String");
    chunk.emit_call(stringify, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, message_slot, line);

    // emit_store_constructor_with_upvalues publishes this exact class name;
    // core constructors have arity 3 and takes_receiver=false. Do not use a JS
    // __ctor_<name> anchor or inject a receiver into this PHP constructor ABI.
    vybe_compiler::primitives::globals::emit_read(chunk, exc_name, line);
    chunk.emit_op_u16(Op::LOCAL_GET, message_slot, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(chunk, 3, line);
    vybe_compiler::primitives::errors::emit_throw(chunk, line);
}

/// Push `is_collection(slot)` as an i32 bool — true when the value is a PHP
/// array/collection or object, false only for true scalars (string/number/
/// bool/null). A PHP array is either a JS **Array** (sequential →
/// `ecma:array.isArray`) or a JS **object** (associative `Map`/`Ordinary` →
/// `recipes::is_object`, `REF_TEST "object"`); neither test alone covers both,
/// so we OR them. This mirrors PHP's own array-vs-scalar boundary and avoids
/// wrongly rejecting either array shape.
fn push_is_collection(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    vybe_compiler::primitives::instructions::recipes::is_object(chunk, line);
    // REF_TEST already leaves an i32 Boolean. isArray has a Boolean result
    // contract, so unbox that result directly rather than emitting the
    // general dynamic truthiness ladder for either predicate.
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    let idx = chunk.add_import("ecma:array", "isArray");
    chunk.emit_call(idx, 1, line);
    let cast = chunk.add_import("wasm:js-boolean", "cast");
    chunk.emit_call(cast, 1, line);
    chunk.emit_op(Op::I32_OR, line);
}

/// Guard the argument at `arg_idx` (0-based) of a builtin whose `argc`
/// arguments are on the operand stack (last argument on top). Throws
/// `exc_name(msg)` when the argument violates `expect`; otherwise leaves all
/// arguments on the stack unchanged for the underlying builtin.
pub fn guard_arg(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    arg_idx: u8,
    expect: Expect,
    exc_name: &str,
    msg: &str,
    line: u32,
) {
    let chunk = &mut chunks[current];
    // Values below the guarded argument can remain on the operand stack.
    // Keep the guarded value there too: local.tee gives the predicate a
    // stable local without requiring another reload on the valid path.
    // Invalid/missing-argument shapes retain their previous emission path.
    let keep_guarded = arg_idx < argc;
    let first = if keep_guarded { arg_idx } else { 0 };
    let stored = argc - first;
    let base = chunk.local_count;
    chunk.local_count += u16::from(stored);
    for i in (first..argc).rev() {
        let op = if keep_guarded && i == arg_idx {
            Op::LOCAL_TEE
        } else {
            Op::LOCAL_SET
        };
        chunk.emit_op_u16(op, base + u16::from(i - first), line);
    }

    let slot = base + u16::from(arg_idx - first);
    match expect {
        // PHP arrays exclude ordinary objects and ArrayObject instances.
        Expect::Array => {
            super::copy_adapter::emit_php_arrayish_slot(chunk, slot, line);
            chunk.emit_op(Op::I32_EQZ, line);
        }
        Expect::NotScalar => {
            // bad = !is_collection
            push_is_collection(chunk, slot, line);
            chunk.emit_op(Op::I32_EQZ, line);
        }
        Expect::StringCoercible => {
            push_is_collection(chunk, slot, line);
            chunk.emit_if_i32(line);
            let method = class_slots::resolve(
                &class_slots::ClassSlot::Slot(vybe_ast::ProtocolSlot::ToString),
                &class_slots::PlainNames,
            );
            class_slots::emit_class_get(
                chunk,
                class_slots::ObjSource::Local(slot),
                &method,
                class_slots::Dest::Stack,
                line,
            );
            chunk.emit_op(Op::REF_IS_NULL, line);
            chunk.emit_else(line);
            chunk.emit_i32_const(0, line);
            chunk.emit_end(line);
        }
        // Accept a scalar; an array/object is illegal.
        Expect::NotArray => {
            // bad = is_collection
            push_is_collection(chunk, slot, line);
        }
    }
    chunk.emit_if(line);
    let _ = chunk;
    emit_throw_const(chunks, current, exc_name, msg, line);
    let chunk = &mut chunks[current];
    chunk.emit_end(line);

    // The prefix and guarded value are already in their original positions.
    // Restore only arguments that were above the checked value.
    for i in u8::from(keep_guarded)..stored {
        chunk.emit_op_u16(Op::LOCAL_GET, base + u16::from(i), line);
    }
}

#[cfg(test)]
#[path = "type_guard_emission_tests.rs"]
mod type_guard_emission_tests;
