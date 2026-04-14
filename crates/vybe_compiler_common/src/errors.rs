//! Exception handling helpers — shared try/catch/finally bytecode patterns.
//!
//! All compilers emit the same opcodes for exception handling:
//! - try_start → body → try_end → handler
//! - try_table for typed multi-catch

use std::sync::Arc;
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

/// Build a standard exception constructor chunk.
/// All languages should use this shape: { __type, __exception_type, name, message }.
/// This ensures Python `except ValueError` can catch a Dart `throw ValueError("...")`.
pub fn emit_exception_constructor(chunk: &mut Chunk, this_slot: u16, exc_name: &str, msg_slot: u16, line: u32) {
    // Create object
    chunk.emit_op_u16(Op::struct_new, 0, line);
    chunk.emit_op_u16(Op::local_set, this_slot, line);
    chunk.emit_op(Op::drop, line);

    // __type = exc_name (for ref_test matching)
    chunk.emit_op_u16(Op::local_get, this_slot, line);
    let t_val = chunk.add_constant(Value::String(Arc::from(exc_name)));
    chunk.emit_op_u16(Op::r#const, t_val, line);
    let t_key = chunk.add_constant(Value::String(Arc::from("__type")));
    chunk.emit_op_u16(Op::struct_set, t_key, line);
    chunk.emit_op(Op::drop, line);

    // __exception_type = exc_name (Python convention)
    chunk.emit_op_u16(Op::local_get, this_slot, line);
    let et_val = chunk.add_constant(Value::String(Arc::from(exc_name)));
    chunk.emit_op_u16(Op::r#const, et_val, line);
    let et_key = chunk.add_constant(Value::String(Arc::from("__exception_type")));
    chunk.emit_op_u16(Op::struct_set, et_key, line);
    chunk.emit_op(Op::drop, line);

    // name = exc_name (JS Error convention)
    chunk.emit_op_u16(Op::local_get, this_slot, line);
    let n_val = chunk.add_constant(Value::String(Arc::from(exc_name)));
    chunk.emit_op_u16(Op::r#const, n_val, line);
    let n_key = chunk.add_constant(Value::String(Arc::from("name")));
    chunk.emit_op_u16(Op::struct_set, n_key, line);
    chunk.emit_op(Op::drop, line);

    // message = msg_slot
    chunk.emit_op_u16(Op::local_get, this_slot, line);
    chunk.emit_op_u16(Op::local_get, msg_slot, line);
    let m_key = chunk.add_constant(Value::String(Arc::from("message")));
    chunk.emit_op_u16(Op::struct_set, m_key, line);
    chunk.emit_op(Op::drop, line);
}

/// Standard exception type names shared across all languages.
/// Maps language-specific names to a canonical set.
pub fn canonical_exception_name(name: &str) -> &str {
    // Defensive: walkers occasionally include trailing whitespace from the
    // type span (e.g. C# `catch (Exception e)` produces "Exception "). Trim
    // before matching, otherwise the lowercase compare misses.
    match name.trim().to_lowercase().as_str() {
        // Python → canonical
        "valueerror" | "formaterror" | "formatexception" => "ValueError",
        "typeerror" => "TypeError",
        "keyerror" | "keynotfoundexception" => "KeyError",
        "indexerror" | "indexoutofrangeexception" | "rangerror" => "IndexError",
        "runtimeerror" | "runtimeexception" => "RuntimeError",
        "stopiteration" | "stateexception" => "StopIteration",
        "attributeerror" | "nosuchmethoderror" => "AttributeError",
        "zerodivisionerror" | "integerdivisionbyzeroexception" => "ZeroDivisionError",
        "filenotfounderror" | "filenotfoundexception" => "FileNotFoundError",
        "importerror" => "ImportError",
        "notimplementederror" | "unimplementederror" => "NotImplementedError",
        "overflowerror" | "overflowexception" | "stackoverflowerror" => "OverflowError",
        "ioerror" | "ioexception" => "IOError",
        "oserror" => "OSError",
        "exception" | "error" => "Exception",
        _ => name,
    }
}

/// Emit the start of a try block. Returns the catch_jump offset to patch later.
/// Layout: [try_start, u16 catch_offset, u16 finally_offset]
/// Stack: unchanged
pub fn emit_try_start(chunk: &mut Chunk, line: u32) -> usize {
    let catch_jump = chunk.emit_jump(Op::try_start, line);
    chunk.emit(0u8, line); // finally offset high byte (reserved)
    chunk.emit(0u8, line); // finally offset low byte (reserved)
    catch_jump
}

/// Emit the end of the try body (normal exit path).
pub fn emit_try_end(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::try_end, line);
}

/// Patch the catch handler offset after the handler code has been emitted.
///
/// The VM reads `catch_offset` then `finally_offset` (2+2 bytes) before computing
/// `catch_ip = ip + catch_offset`, where ip is *after* all 4 operand bytes.
/// `chunk.patch_jump` assumes ip is right after the 2 patched bytes, so we
/// subtract 2 to account for the extra finally-offset bytes the VM skips.
pub fn patch_catch(chunk: &mut Chunk, catch_jump: usize) {
    let jump = chunk.current_offset() as i32 - (catch_jump as i32 + 2) - 2;
    chunk.code[catch_jump] = (jump >> 8) as u8;
    chunk.code[catch_jump + 1] = (jump & 0xff) as u8;
}

/// Emit a throw — takes the exception value from TOS.
/// Stack before: [exception_value]  Stack after: diverges
pub fn emit_throw(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::throw, line);
}

/// Returns true if `name` (case-insensitive) is one of the known
/// exception type names that should produce the canonical 4-field
/// shape via `emit_exception_new`. The list is the union of every
/// language's built-in exception types — adding a new language entry
/// only requires extending `canonical_exception_name`.
pub fn is_exception_type(name: &str) -> bool {
    let lower = name.to_lowercase();
    matches!(lower.as_str(),
        // Generic
        "exception" | "error" | "throwable"
        // Python / canonical
        | "valueerror" | "typeerror" | "keyerror" | "indexerror"
        | "runtimeerror" | "stopiteration" | "attributeerror"
        | "zerodivisionerror" | "filenotfounderror" | "importerror"
        | "notimplementederror" | "overflowerror" | "ioerror" | "oserror"
        // .NET / VB / C#
        | "systemexception" | "argumentexception" | "argumentnullexception"
        | "invalidoperationexception" | "notimplementedexception"
        | "notsupportedexception" | "nullreferenceexception"
        | "indexoutofrangeexception" | "keynotfoundexception"
        | "formatexception" | "stackoverflowerror" | "stackoverflowexception"
        | "integerdivisionbyzeroexception" | "rangerror" | "stateexception"
        | "filenotfoundexception" | "ioexception" | "formaterror"
        | "nosuchmethoderror" | "unimplementederror" | "overflowexception"
        // PHP
        | "runtimeexception" | "logicexception" | "domainexception"
        | "lengthexception" | "outofboundsexception" | "outofrangeexception"
        | "rangeexception" | "underflowexception"
        | "unexpectedvalueexception"
        // JS
        | "rangeerror" | "syntaxerror" | "referenceerror" | "urierror"
        // Ruby
        | "standarderror" | "argumenterror" | "nameerror" | "nomethoderror"
    )
}

/// Stack-based exception constructor. Use this in two phases:
///
/// 1. Caller emits `Op::struct_new` and `Op::dup` to push `[obj, obj]`,
///    then emits the message expression to push `[obj, obj, msg]`.
/// 2. Caller invokes `emit_exception_new_finalize(chunk, exc_name, line)`
///    which consumes the inner `[obj, msg]` pair into `obj.message=msg`,
///    then stamps `__type`, `__exception_type`, `name` onto the outer obj.
///
/// Stack before: `[obj, obj, msg]`   Stack after: `[obj]`
///
/// Splitting the helper this way avoids the closure-vs-`&mut self`
/// borrow problem in language compilers (the compiler needs `&mut self`
/// to emit the message expression, which can't co-exist with a `&mut
/// chunk` borrow held by a closure-taking helper).
///
/// This is the **single source of truth** for `new SomeError(msg)` across
/// every language compiler. The name is normalized via
/// `canonical_exception_name` so PHP `RuntimeException`, Python
/// `RuntimeError`, JS `Error`, etc. all produce identical bytecode and
/// can therefore catch each other across language boundaries.
pub fn emit_exception_new_finalize(chunk: &mut Chunk, exc_name: &str, line: u32) {
    let canon = canonical_exception_name(exc_name);
    // Preserve the original name for the `name` property (JS expects
    // `Error`, `RangeError`, etc. — not the canonical cross-language form).
    let original = exc_name.trim();

    // [obj, obj, msg] → [obj, msg_val] via struct_set "message"
    let m_key = chunk.add_constant(Value::String(Arc::from("message")));
    chunk.emit_op_u16(Op::struct_set, m_key, line);
    // [obj, msg_val] → [obj]
    chunk.emit_op(Op::drop, line);

    // __type and __exception_type use the canonical name (for cross-language
    // catch dispatch). `name` uses the original (language-specific) name
    // so `err.name` returns what the language expects.
    for (key, val) in [("__type", canon), ("__exception_type", canon), ("name", original)] {
        chunk.emit_op(Op::dup, line);
        let v = chunk.add_constant(Value::String(Arc::from(val)));
        chunk.emit_op_u16(Op::r#const, v, line);
        let k = chunk.add_constant(Value::String(Arc::from(key)));
        chunk.emit_op_u16(Op::struct_set, k, line);
        chunk.emit_op(Op::drop, line);
    }

    // `stack` = "Name: message" for JS Error.stack compatibility.
    // Stack: [obj]. Read message back, concat with prefix, stamp.
    // `stack` = "Name: message" for JS Error.stack compatibility.
    // Build from message (already on the object) by emitting the caller
    // to do it — see the `stamp_error_stack` parameter pattern below.
    // For now, stamp the `stack` field in the vybex compiler where we
    // have access to local variables for the swap.
}

/// Emit type-dispatch for a single catch arm. The exception object is
/// expected on TOS (and remains on TOS for the next arm to test).
/// Returns the patch offset of the "skip-this-arm" jump that the caller
/// must patch after the arm body has been emitted.
///
/// `expected_canon` is the canonical exception type name (already passed
/// through `canonical_exception_name`). The literal string `"Exception"`
/// matches anything (catch-all).
///
/// Stack before: [exc]  Stack after: [exc] (preserved for next arm)
pub fn emit_catch_dispatch(chunk: &mut Chunk, expected_canon: &str, line: u32) -> usize {
    if expected_canon == "Exception" || expected_canon.is_empty() {
        // Catch-all — no dispatch, no skip jump.
        // Caller can patch a no-op offset, so return current_offset.
        return usize::MAX;
    }
    // dup exc, struct_get __exception_type, push expected, dyn_eq, br_if_false skip
    chunk.emit_op(Op::dup, line);
    let k = chunk.add_constant(Value::String(Arc::from("__exception_type")));
    chunk.emit_op_u16(Op::struct_get, k, line);
    let v = chunk.add_constant(Value::String(Arc::from(expected_canon)));
    chunk.emit_op_u16(Op::r#const, v, line);
    chunk.emit_op(Op::dyn_eq, line);
    chunk.emit_jump(Op::br_if_false, line)
}
