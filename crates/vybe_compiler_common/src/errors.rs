//! Exception handling helpers — shared try/catch/finally bytecode patterns.
//!
//! All compilers emit the same opcodes for exception handling:
//! - try_start → body → try_end → handler
//! - try_table for typed multi-catch

use std::rc::Rc;
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
    let t_val = chunk.add_constant(Value::String(Rc::from(exc_name)));
    chunk.emit_op_u16(Op::r#const, t_val, line);
    let t_key = chunk.add_constant(Value::String(Rc::from("__type")));
    chunk.emit_op_u16(Op::struct_set, t_key, line);
    chunk.emit_op(Op::drop, line);

    // __exception_type = exc_name (Python convention)
    chunk.emit_op_u16(Op::local_get, this_slot, line);
    let et_val = chunk.add_constant(Value::String(Rc::from(exc_name)));
    chunk.emit_op_u16(Op::r#const, et_val, line);
    let et_key = chunk.add_constant(Value::String(Rc::from("__exception_type")));
    chunk.emit_op_u16(Op::struct_set, et_key, line);
    chunk.emit_op(Op::drop, line);

    // name = exc_name (JS Error convention)
    chunk.emit_op_u16(Op::local_get, this_slot, line);
    let n_val = chunk.add_constant(Value::String(Rc::from(exc_name)));
    chunk.emit_op_u16(Op::r#const, n_val, line);
    let n_key = chunk.add_constant(Value::String(Rc::from("name")));
    chunk.emit_op_u16(Op::struct_set, n_key, line);
    chunk.emit_op(Op::drop, line);

    // message = msg_slot
    chunk.emit_op_u16(Op::local_get, this_slot, line);
    chunk.emit_op_u16(Op::local_get, msg_slot, line);
    let m_key = chunk.add_constant(Value::String(Rc::from("message")));
    chunk.emit_op_u16(Op::struct_set, m_key, line);
    chunk.emit_op(Op::drop, line);
}

/// Standard exception type names shared across all languages.
/// Maps language-specific names to a canonical set.
pub fn canonical_exception_name(name: &str) -> &str {
    match name.to_lowercase().as_str() {
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
