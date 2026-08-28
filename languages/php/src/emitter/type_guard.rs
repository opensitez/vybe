//! Runtime argument type-guards for PHP builtins.
//!
//! PHP raises `TypeError`/`ValueError` where JavaScript would silently coerce
//! or return `undefined` (and where Vybe would otherwise crash trying to treat
//! a scalar as an array). Each guard is emitted at a builtin's dispatch site:
//! it inspects an argument already on the operand stack and, if the type is
//! illegal, throws through the **common** errors emitter
//! (`errors::emit_exception_new_finalize`) so the thrown value is byte-identical
//! to — and catchable across — every other language (PHP `TypeError` ≡ JS
//! `TypeError` ≡ Python `TypeError`).
//!
//! The guard spills the `argc` arguments to locals, tests the target argument,
//! throws on failure, then restores every argument in its original order so the
//! underlying builtin emit consumes an unchanged stack.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;
use vybe_compiler::primitives::class_slots::{
    self,
};

/// What an argument is required to be for a builtin to accept it.
#[derive(Clone, Copy)]
pub enum Expect {
    /// Must be a PHP array; a non-array (scalar or object) is illegal.
    /// (`array_merge`, `sort`, `array_keys`, `in_array` haystack, …)
    Array,
    /// Must be coercible to string/number; an array is illegal but scalars are
    /// fine (`strlen`, `str_starts_with`, `abs`, `round`).
    NotArray,
    /// Must be a collection or object; a bare scalar is illegal but objects
    /// (Countable, Traversable, generators) are fine (`count`,
    /// `iterator_to_array`).
    NotScalar,
}

/// The PHP `Throwable` ancestry for a built-in exception, most-derived first.
/// `Error` and `Exception` are **sibling** branches of `Throwable` (unlike JS,
/// where `Error` is the universal base), and `TypeError`/`ValueError` extend
/// `Error`. The catch-matcher (`primitives/mod.rs`) matches a caught type against
/// the exception's `__types` array, so stamping the full chain is what lets
/// `catch (\Error)` catch a thrown `TypeError` — mirroring the chain that user
/// `throw new TypeError()` gets from class emission.
fn php_exception_chain(exc_name: &str) -> &'static [&'static str] {
    match exc_name {
        "TypeError" => &["TypeError", "Error", "Throwable"],
        "ValueError" => &["ValueError", "Error", "Throwable"],
        "ArgumentCountError" => &["ArgumentCountError", "TypeError", "Error", "Throwable"],
        "DivisionByZeroError" => &[
            "DivisionByZeroError",
            "ArithmeticError",
            "Error",
            "Throwable",
        ],
        "FiberError" => &["FiberError", "Error", "Throwable"],
        "Error" => &["Error", "Throwable"],
        "Exception" => &["Exception", "Throwable"],
        "PDOException" => &["PDOException", "Exception", "Throwable"],
        "RuntimeException" => &["RuntimeException", "Exception", "Throwable"],
        "OutOfRangeException" => &[
            "OutOfRangeException",
            "LogicException",
            "Exception",
            "Throwable",
        ],
        "OutOfBoundsException" => &[
            "OutOfBoundsException",
            "RuntimeException",
            "Exception",
            "Throwable",
        ],
        _ => &[],
    }
}

/// Unconditionally construct and throw `exc_name(msg)` at the current point.
/// Diverges. Uses the shared errors emitter to mint the instance and the shared
/// `emit_instanceof_chain` to stamp its `Throwable` ancestry, so the exception
/// participates in cross-language `catch` matching including base-class catches
/// (`catch (\Error)` matching a `TypeError`).
pub fn emit_throw_const(
    chunks: &mut [Chunk],
    current: usize,
    exc_name: &str,
    msg: &str,
    line: u32,
) {
    let chunk = &mut chunks[current];
    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_dup(line);
    chunk.emit_string_const(msg, line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(chunk, exc_name, line);
    // Park the instance in a local so the shared chain-stamper can push each
    // ancestor name into its `__types` array, then reload and throw.
    let this_slot = chunk.local_count;
    chunk.local_count += 1;
    chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);
    for name in php_exception_chain(exc_name) {
        vybe_compiler::primitives::reflection::emit_instanceof_chain(
            chunks, current, this_slot, name, line,
        );
    }
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
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
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    let idx = chunk.add_import("ecma:array", "isArray");
    chunk.emit_call(idx, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
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
    // Reserve a contiguous block of locals: arg0 → base, argN-1 → base+argc-1.
    let base = chunk.local_count;
    chunk.local_count += argc as u16;
    for i in (0..argc).rev() {
        chunk.emit_op_u16(Op::LOCAL_SET, base + i as u16, line);
    }

    let slot = base + arg_idx as u16;
    match expect {
        // Accept a PHP array/collection (JS array or object); a scalar is illegal.
        Expect::Array | Expect::NotScalar => {
            // bad = !is_collection
            push_is_collection(chunk, slot, line);
            chunk.emit_op(Op::I32_EQZ, line);
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

    // Restore arguments in original order.
    for i in 0..argc {
        chunk.emit_op_u16(Op::LOCAL_GET, base + i as u16, line);
    }
}
