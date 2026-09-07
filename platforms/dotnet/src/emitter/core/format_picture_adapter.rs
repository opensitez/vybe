//! VB.NET / VBA `Format(value, picture)` picture-string adapter.
//!
//! Distinct from C#'s `String.Format("{0}", value)` (handled by
//! `string_format_adapter`): VB's `Format(value, picture)` accepts a
//! picture string like `"0.00"` / `"#,##0.00"` / `"$#,##0.00"` /
//! `"yyyy-MM-dd"` and applies the rendering rules per .NET's
//! `IFormatProvider` numeric/date format specifiers.
//!
//! This minimal adapter covers the common numeric pictures by
//! detecting digit-pattern grammar at runtime:
//!
//!   `"0"`         → integer rendering, no decimals
//!   `"0.0"` …     → fixed N decimal places (count of `0`s after `.`)
//!   `"#,##0"`     → integer with thousands separators
//!   `"#,##0.00"`  → thousands separators + N decimals
//!   `"$<picture>"` → currency prefix + numeric picture
//!
//! Anything else falls back to `String(value)`. Date pictures
//! (`"yyyy-MM-dd"` etc.) currently fall through — bring them back as a
//! follow-up adapter (`format_picture_date_adapter`) if a test demands.
//!
//! Stack on entry: `[value, picture]` (argc=2) or `[value]` (argc=1,
//! picture defaults to `""` → toString). Stack on exit: `[string]`.

use std::sync::Arc;
use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource};
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use super::object_fields::field_slot;

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        _ => panic!("push_const: no WASM-compliant encoding for {:?}", val),
    }
}

/// Emit `Format(value, [picture])` at the call site.
pub fn emit_format_picture(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc == 0 {
        push_const(chunk, Value::String(Arc::from("")), line);
        return;
    }
    if argc == 1 {
        // Bare value — toString and return.
        let to_str = chunk.add_import("ecma:string", "String");
        chunk.emit_call(to_str, 1, line);
        return;
    }

    // Two-arg path: [value, picture].
    // Trailing args beyond 2 are dropped (defensive — VB Format is
    // strictly 1-or-2-arity in practice).
    for _ in 2..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let value_slot = chunk.alloc_scratch(2);
    let picture_slot = value_slot + 1;
    // Stash picture (top), then value.
    chunk.emit_op_u16(Op::LOCAL_SET, picture_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Stack,
        &field_slot("__time"),
        Dest::Stack,
        line,
    );
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, picture_slot, line);
    emit_vb_date_picture_alias(chunk, line);
    super::datetime_format_adapter::emit_date_format(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_else(line);

    vybe_compiler::primitives::globals::emit_read(chunk, "__vybe_vb_format", line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, picture_slot, line);
    chunk.emit_op_u8_u8(Op::CALL_REF, 2, 1, line);
    chunk.emit_end(line);
}

fn emit_vb_date_picture_alias(chunk: &mut Chunk, line: u32) {
    let picture_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, picture_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, picture_slot, line);
    chunk.emit_string_const("Short Date", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("d", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, picture_slot, line);
    chunk.emit_string_const("Short Time", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("t", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, picture_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}
