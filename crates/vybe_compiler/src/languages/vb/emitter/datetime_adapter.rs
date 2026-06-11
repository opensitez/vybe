//! VB date/time builtin adapters — bytecode-only.
//!
//! VB builtins like `Year(d)`, `Month(d)`, `Day(d)`, `Hour(d)`, `Minute(d)`,
//! `Second(d)`, and `Weekday(d)` extract fields from a DateTime object
//! produced by `emit_datetime_now` / `emit_datetime_today`.  The DateTime
//! struct already carries all these as named fields, so the adapters are
//! plain STRUCT_GET calls — no host imports needed.

use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

fn struct_get_field(chunk: &mut Chunk, field: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(field)));
    chunk.emit_op_u16(Op::STRUCT_GET, idx, line);
}

/// `Year(d)` — extracts the Year field from a DateTime object.
pub fn emit_vb_year(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    struct_get_field(&mut chunks[current], "Year", line);
}

/// `Month(d)` — extracts the Month field (1–12).
pub fn emit_vb_month(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    struct_get_field(&mut chunks[current], "Month", line);
}

/// `Day(d)` — extracts the Day field (1–31).
pub fn emit_vb_day(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    struct_get_field(&mut chunks[current], "Day", line);
}

/// `Hour(d)` — extracts the Hour field (0–23).
pub fn emit_vb_hour(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    struct_get_field(&mut chunks[current], "Hour", line);
}

/// `Minute(d)` — extracts the Minute field (0–59).
pub fn emit_vb_minute(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    struct_get_field(&mut chunks[current], "Minute", line);
}

/// `Second(d)` — extracts the Second field (0–59).
pub fn emit_vb_second(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    struct_get_field(&mut chunks[current], "Second", line);
}

/// `Weekday(d)` — extracts the DayOfWeek string field.
/// VB returns 1 (Sunday) through 7 (Saturday); here we return the
/// string name. Callers that need the integer can use WeekdayName.
pub fn emit_vb_weekday(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    struct_get_field(&mut chunks[current], "DayOfWeek", line);
}
