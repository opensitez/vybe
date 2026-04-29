//! .NET `System.DateTime` adapter — bytecode-only.
//!
//! `DateTime` is a calendar instant; .NET's `DateTime.Now /
//! UtcNow / Parse / Today` static methods plus the `New
//! DateTime(y, m, d, ...)` constructor produce a value with
//! `__type=DateTime` and a millisecond-since-epoch payload.
//!
//! The underlying primitive is `wasi:clocks/wall-clock.now` (WASI
//! 0.2.11 spec primitive — registered alongside the legacy flat
//! `wasi:clocks` namespace). `ecma:date.now` reads through it and
//! returns ms since epoch — the form ECMA-262 §21.4 [[DateValue]]
//! uses. Each adapter wraps that ms in a DateTime-shaped Object so
//! the .NET surface looks .NET-shaped while the bytecode is
//! standardized.

use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;
use std::sync::Arc;

const TYPE_KEY: &str = "__type";
const TIME_KEY: &str = "__time";

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    let idx = chunk.add_constant(val);
    chunk.emit_op_u16(Op::CONST, idx, line);
}

/// Wrap a millisecond timestamp on stack-top as a DateTime object.
/// Stack on entry: `[ms]` ; Stack on exit: `[datetime_obj]`.
fn emit_wrap_ms(chunk: &mut Chunk, line: u32) {
    let type_key = chunk.add_constant(Value::String(Arc::from(TYPE_KEY)));
    let time_key = chunk.add_constant(Value::String(Arc::from(TIME_KEY)));
    let ms_slot = chunk.local_count;
    chunk.local_count = ms_slot + 1;

    chunk.emit_op_u16(Op::LOCAL_SET, ms_slot, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op(Op::DUP, line);
    push_const(chunk, Value::String(Arc::from("DateTime")), line);
    chunk.emit_op_u16(Op::STRUCT_SET, type_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, time_key, line);
    chunk.emit_op(Op::DROP, line);
}

/// `DateTime.Now` / `DateTime.UtcNow` — read `ecma:date.now` (which
/// reads through `wasi:clocks/wall-clock.now`) and wrap in a
/// DateTime object.
///
/// Stack on entry: `[]` ; Stack on exit: `[datetime_obj]`
pub fn emit_datetime_now(chunks: &mut [Chunk], current: usize, line: u32) {
    let now_idx = chunks[0].add_import("ecma:date", "now");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, now_idx, line);
    chunk.emit(0, line);
    emit_wrap_ms(chunk, line);
}

/// `DateTime.Parse(s)` — parse a date string via `ecma:date.parse`
/// (ECMA-262 §21.4.3.2) and wrap.
///
/// Stack on entry: `[s]` ; Stack on exit: `[datetime_obj]`
pub fn emit_datetime_parse(chunks: &mut [Chunk], current: usize, line: u32) {
    let parse_idx = chunks[0].add_import("ecma:date", "parse");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, parse_idx, line);
    chunk.emit(1, line);
    emit_wrap_ms(chunk, line);
}

/// `DateTime.Today` — synonym for `Now` in .NET (returns midnight of
/// today; we return the current instant for the MVP). Same bytecode
/// as `Now`.
pub fn emit_datetime_today(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_datetime_now(chunks, current, line);
}
