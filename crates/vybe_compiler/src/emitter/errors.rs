//! Exception handling helpers — shared try/catch/finally bytecode patterns.
//!
//! All compilers emit the same opcodes for exception handling:
//! - try_table (real WASM EH Phase 4) → body → try_end → handler
//! - try_end pops the handler on normal (non-throwing) exit

use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

/// Build a standard exception constructor chunk.
/// All languages should use this shape: { __type, __exception_type, name, message }.
/// This ensures Python `except ValueError` can catch a Dart `throw ValueError("...")`.
pub fn emit_exception_constructor(
    chunk: &mut Chunk,
    this_slot: u16,
    exc_name: &str,
    msg_slot: u16,
    line: u32,
) {
    // Create object
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);

    // __type = exc_name (for ref_test matching)
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_string_const(exc_name, line);
    let t_key = chunk.add_constant(Value::String(Arc::from("__type")));
    chunk.emit_op_u16(Op::STRUCT_SET, t_key, line);
    chunk.emit_op(Op::DROP, line);

    // __exception_type = exc_name (Python convention)
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_string_const(exc_name, line);
    let et_key = chunk.add_constant(Value::String(Arc::from("__exception_type")));
    chunk.emit_op_u16(Op::STRUCT_SET, et_key, line);
    chunk.emit_op(Op::DROP, line);

    // name = exc_name (JS Error convention)
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_string_const(exc_name, line);
    let n_key = chunk.add_constant(Value::String(Arc::from("name")));
    chunk.emit_op_u16(Op::STRUCT_SET, n_key, line);
    chunk.emit_op(Op::DROP, line);

    // message = msg_slot
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, msg_slot, line);
    let m_key = chunk.add_constant(Value::String(Arc::from("message")));
    chunk.emit_op_u16(Op::STRUCT_SET, m_key, line);
    chunk.emit_op(Op::DROP, line);
}

/// Standard exception type names shared across all languages.
/// Maps language-specific names to a canonical set.
pub fn canonical_exception_name(name: &str) -> &str {
    // Defensive: walkers occasionally include trailing whitespace from the
    // type span (e.g. C# `catch (Exception e)` produces "Exception "). Trim
    // before matching AND in the fallthrough so the runtime-side
    // `STRUCT_GET __exception_type` compare doesn't miss on a trailing
    // space mismatch.
    let trimmed = name.trim();
    let short_name = trimmed.rsplit('.').next().unwrap_or(trimmed).trim();
    match short_name.to_lowercase().as_str() {
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
        _ => trimmed,
    }
}

/// Emit the start of a try block. Returns the offset_pos to patch later.
/// Layout: [try_table, u8 handler_count=1, u8 tag=0, u16 catch_offset]
/// Stack: unchanged
pub fn emit_try_start(chunk: &mut Chunk, line: u32) -> usize {
    chunk.emit_op(Op::TRY_TABLE, line); // real WASM Phase 4 EH opcode
    chunk.emit(1u8, line); // handler_count = 1
    chunk.emit(0u8, line); // tag = 0 (catch-all)
    let offset_pos = chunk.current_offset();
    chunk.emit(0u8, line); // catch offset hi (placeholder)
    chunk.emit(0u8, line); // catch offset lo (placeholder)
    offset_pos
}

/// Emit the end of the try body (normal exit path).
pub fn emit_try_end(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::END, line);
}

/// Patch the catch handler offset after the handler code has been emitted.
///
/// The VM reads `offset` (2 bytes) and computes `catch_ip = ip + offset`,
/// where ip is the position right after those 2 bytes (`offset_pos + 2`).
/// The forward distance from that ip to the current end of code is the offset.
pub fn patch_catch(chunk: &mut Chunk, offset_pos: usize) {
    let jump = chunk.current_offset() as i32 - (offset_pos as i32 + 2);
    chunk.code[offset_pos] = (jump >> 8) as u8;
    chunk.code[offset_pos + 1] = (jump & 0xff) as u8;
}

/// Emit a throw — takes the exception value from TOS.
/// Stack before: [exception_value]  Stack after: diverges
pub fn emit_throw(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::THROW, line);
}

/// Returns true if `name` (case-insensitive) is one of the known
/// exception type names that should produce the canonical 4-field
/// shape via `emit_exception_new`. The list is the union of every
/// language's built-in exception types — adding a new language entry
/// only requires extending `canonical_exception_name`.
pub fn is_exception_type(name: &str) -> bool {
    let lower = name.to_lowercase();
    matches!(
        lower.as_str(),
        // Generic
        "exception" | "error" | "throwable"
        // Python / canonical
        | "valueerror" | "typeerror" | "keyerror" | "indexerror"
        | "runtimeerror" | "stopiteration" | "attributeerror"
        | "zerodivisionerror" | "filenotfounderror" | "importerror"
        | "notimplementederror" | "overflowerror" | "ioerror" | "oserror"
        // .NET / VB / C#
        | "systemexception" | "applicationexception" | "argumentexception" | "argumentnullexception"
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
        | "unhandledmatcherror" | "divisionbyzeroerror" | "argumentcounterror"
        | "errorexception"
        // JS
        | "rangeerror" | "syntaxerror" | "referenceerror" | "urierror"
        | "evalerror" | "aggregateerror"
        // Ruby
        | "standarderror" | "argumenterror" | "nameerror" | "nomethoderror"
    )
}

/// Stack-based exception constructor. Use this in two phases:
///
/// 1. Caller emits `Op::STRUCT_NEW` and `Op::DUP` to push `[obj, obj]`,
///    then emits the message expression to push `[obj, obj, msg]`.
/// 2. Caller invokes `emit_exception_new_finalize(chunk, exc_name, line)`
///    which consumes the inner `[obj, msg]` pair into `obj.message=msg`,
///    then stamps `__type`, `__exception_type` onto the outer obj.
///
/// Per ECMA-262 §20.5, name and constructor are inherited from Error.prototype,
/// not own properties, so they are not set here. JavaScript callers should ensure
/// proper prototype chain setup if needed.
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
    let original = exc_name.trim();

    // [obj, obj, msg] → [obj, msg_val] via struct_set "message"
    let m_key = chunk.add_constant(Value::String(Arc::from("message")));
    chunk.emit_op_u16(Op::STRUCT_SET, m_key, line);
    // [obj, msg_val] → [obj]
    chunk.emit_op(Op::DROP, line);

    // __type and __exception_type use the canonical name (for cross-language
    // catch dispatch).
    for (key, val) in [("__type", canon), ("__exception_type", canon)] {
        chunk.emit_dup(line);
        chunk.emit_string_const(val, line);
        let k = chunk.add_constant(Value::String(Arc::from(key)));
        chunk.emit_op_u16(Op::STRUCT_SET, k, line);
        chunk.emit_op(Op::DROP, line);
    }

    // Set name as a dynamic (non-indexed) property with the original language-specific name.
    // It will be added to __nonenum at the type level, making it non-enumerable.
    chunk.emit_dup(line);
    chunk.emit_string_const(original, line);
    let n_key = chunk.add_constant(Value::String(Arc::from("name")));
    chunk.emit_op_u16(Op::STRUCT_SET, n_key, line);
    chunk.emit_op(Op::DROP, line);
}

/// Emit the disposal half of a resource-management block (C# `using`,
/// Python `with`, Java try-with-resources, JS `using x = …`). Reads
/// the resource from `slot` and calls its lifecycle method (`Dispose`,
/// `__exit__`, `close`, …) if defined. Guards against the method
/// being absent so resources without a disposer don't trap.
///
/// ECMA-334 §13.14 / Python §8.5 / JS Stage 3 explicit-resource-
/// management share the same lowering: `try { body; } finally {
/// dispose; }`. We emit just the dispose tail; full try/finally
/// wrapping is the caller's job (or future enhancement here).
///
/// `dispose_method`: the canonical method name (`"Dispose"` for .NET,
/// `"__exit__"` for Python, `"close"` for Java AutoCloseable, etc.).
pub fn emit_resource_dispose(chunk: &mut Chunk, slot: u16, dispose_method: &str, line: u32) {
    let dispose_key = chunk.add_constant(Value::String(Arc::from(dispose_method)));
    let dispose_block = chunk.emit_block(line);
    // method = resource[<dispose_method>]
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, dispose_key, line);
    // if method is null/undefined, skip the call.
    chunk.emit_dup(line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_br_if(0, line);
    // Stack: [method]. Push receiver and CALL_REF(1). Drop result.
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    chunk.emit_op(Op::DROP, line);
    // Skipped path leaves `method` (null/undef) on stack — the END
    // closes the block, after which we DROP unconditionally.
    chunk.emit_end(line);
    chunk.patch_block(dispose_block);
    chunk.emit_op(Op::DROP, line);
}
