//! .NET `System.TimeSpan` adapter — bytecode-only.
//!
//! `TimeSpan` is a duration value type; .NET's
//! `TimeSpan.From{Days,Hours,Minutes,Seconds,Milliseconds}(n)`
//! factory methods build a duration record from a unit count.
//! There's no ECMA-262 mirror (JS uses raw `number` ms), but the
//! arithmetic is trivial: multiply by the unit-to-ms factor and
//! stash on a struct.
//!
//! Each adapter emits inline bytecode — no host fns. The result has
//! shape `{ __type: "TimeSpan", totalmilliseconds, totalseconds,
//! totalminutes, totalhours, totaldays, days, hours, minutes,
//! seconds }` matching the existing `vybe:types/timeSpan*` host
//! impls so callers continue to work.

use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;
use std::sync::Arc;

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    let idx = chunk.add_constant(val);
    chunk.emit_op_u16(Op::CONST, idx, line);
}

fn struct_set_field(chunk: &mut Chunk, key_idx: u16, line: u32) {
    chunk.emit_op_u16(Op::STRUCT_SET, key_idx, line);
    chunk.emit_op(Op::DROP, line);
}

/// Build the TimeSpan object given the total milliseconds on the stack.
/// Stack on entry: `[total_ms]` ; Stack on exit: `[ts_obj]`
fn emit_build_timespan(chunk: &mut Chunk, line: u32) {
    let ms_slot = chunk.local_count;
    chunk.local_count = ms_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, ms_slot, line);
    chunk.emit_op(Op::DROP, line);

    // Pre-compute constants for the various unit conversions.
    let type_key = chunk.add_constant(Value::String(Arc::from("__type")));
    let total_ms_key = chunk.add_constant(Value::String(Arc::from("totalmilliseconds")));
    let total_sec_key = chunk.add_constant(Value::String(Arc::from("totalseconds")));
    let total_min_key = chunk.add_constant(Value::String(Arc::from("totalminutes")));
    let total_hr_key = chunk.add_constant(Value::String(Arc::from("totalhours")));
    let total_day_key = chunk.add_constant(Value::String(Arc::from("totaldays")));

    // STRUCT_NEW
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    // [obj]

    // __type = "TimeSpan"
    chunk.emit_op(Op::DUP, line);
    push_const(chunk, Value::String(Arc::from("TimeSpan")), line);
    struct_set_field(chunk, type_key, line);

    // totalmilliseconds = ms
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    struct_set_field(chunk, total_ms_key, line);

    // totalseconds = ms / 1000
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(1000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    struct_set_field(chunk, total_sec_key, line);

    // totalminutes = ms / 60_000
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(60_000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    struct_set_field(chunk, total_min_key, line);

    // totalhours = ms / 3_600_000
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(3_600_000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    struct_set_field(chunk, total_hr_key, line);

    // totaldays = ms / 86_400_000
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(86_400_000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    struct_set_field(chunk, total_day_key, line);
}

/// Build a TimeSpan from a count of `unit_ms` units. Stack: `[n]` →
/// `[ts]`. Internally: `total_ms = n * unit_ms`, then build the
/// record. Generic over unit so all `From*` methods share one body.
fn emit_timespan_from_unit(chunks: &mut [Chunk], current: usize, unit_ms: f64, line: u32) {
    let chunk = &mut chunks[current];
    push_const(chunk, Value::F64(unit_ms), line);
    chunk.emit_op(Op::F64_MUL, line);
    emit_build_timespan(chunk, line);
}

/// `TimeSpan.FromDays(n)` — `n * 86_400_000` ms.
pub fn emit_timespan_from_days(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_timespan_from_unit(chunks, current, 86_400_000.0, line);
}

/// `TimeSpan.FromHours(n)` — `n * 3_600_000` ms.
pub fn emit_timespan_from_hours(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_timespan_from_unit(chunks, current, 3_600_000.0, line);
}

/// `TimeSpan.FromMinutes(n)` — `n * 60_000` ms.
pub fn emit_timespan_from_minutes(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_timespan_from_unit(chunks, current, 60_000.0, line);
}

/// `TimeSpan.FromSeconds(n)` — `n * 1000` ms.
pub fn emit_timespan_from_seconds(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_timespan_from_unit(chunks, current, 1000.0, line);
}

/// `TimeSpan.FromMilliseconds(n)` — pass-through.
pub fn emit_timespan_from_milliseconds(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_timespan_from_unit(chunks, current, 1.0, line);
}

/// `TimeSpan.Zero` — 0-duration TimeSpan. Stack: `[]` → `[ts]`.
pub fn emit_timespan_zero(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    push_const(chunk, Value::F64(0.0), line);
    emit_build_timespan(chunk, line);
}
