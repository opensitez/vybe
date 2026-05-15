//! PHP `$x++` / `$x--` numeric semantics — Rust inline opcode emitter.
//!
//! Mirrors the inline-emit shape used by `datetime_adapter.rs`: each
//! `emit_*(chunks, current, argc, line)` writes WASM opcodes directly
//! into `chunks[current]`. PHP's `++` / `--` are polymorphic — number
//! adds/subtracts 1, string-numeric coerces via `ecma:number.parseFloat`
//! before adding. Pure-alpha bump (`"a"++ → "b"`) is a follow-up.
//!
//! No new host fns; composes only `ecma:number.parseFloat`.

use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;
use std::sync::Arc;

/// `__php_inc(v)` — PHP `$v++` arithmetic.
/// Stack on entry: `[v]` ; Stack on exit: `[v + 1]`.
pub fn emit_php_inc(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_unary_arith(chunks, current, /*plus=*/true, line);
}

/// `__php_dec(v)` — PHP `$v--` arithmetic.
/// Stack on entry: `[v]` ; Stack on exit: `[v - 1]`.
pub fn emit_php_dec(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_unary_arith(chunks, current, /*plus=*/false, line);
}

fn emit_unary_arith(chunks: &mut [Chunk], current: usize, plus: bool, line: u32) {
    let v_slot = {
        let chunk = &mut chunks[current];
        let s = chunk.local_count;
        chunk.local_count = s + 1;
        chunk.emit_op_u16(Op::LOCAL_SET, s, line);
        chunk.emit_op(Op::DROP, line);
        s
    };

    let parse_float = chunks[0].add_import("ecma:number", "parseFloat");

    // typeof(v) === "string"?  if not, BR over the string-coerce arm.
    let chunk = &mut chunks[current];
    let block = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, v_slot, line);
    chunk.emit_op(Op::REF_TYPEOF, line);
    let str_const = chunk.add_constant(Value::String(Arc::from("string")));
    chunk.emit_op_u16(Op::CONST, str_const, line);
    chunk.emit_op(Op::STR_EQUALS, line);
    chunk.emit_op(Op::DYN_NOT, line);
    chunk.emit_br_if(0, line);

    // String case: parseFloat(v) ± 1
    chunk.emit_op_u16(Op::LOCAL_GET, v_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, parse_float, line);
    chunk.emit(1, line);
    let one = chunk.add_constant(Value::F64(1.0));
    chunk.emit_op_u16(Op::CONST, one, line);
    chunk.emit_op(if plus { Op::DYN_ADD } else { Op::F64_SUB }, line);
    let done = chunk.emit_jump(Op::BR, line);
    chunk.emit_end(line);
    chunk.patch_block(block);

    // Numeric case: v ± 1
    chunk.emit_op_u16(Op::LOCAL_GET, v_slot, line);
    let one2 = chunk.add_constant(Value::F64(1.0));
    chunk.emit_op_u16(Op::CONST, one2, line);
    chunk.emit_op(if plus { Op::DYN_ADD } else { Op::F64_SUB }, line);

    chunk.patch_jump(done);
}
