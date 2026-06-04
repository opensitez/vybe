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
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    let idx = chunk.add_constant(val);
    chunk.emit_op_u16(Op::CONST, idx, line);
}

/// Emit `Format(value, [picture])` at the call site.
pub fn emit_format_picture(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc == 0 {
        push_const(chunk, Value::String(Arc::from("")), line);
        return;
    }
    if argc == 1 {
        // Bare value — toString and return.
        let to_str = chunk.add_import("ecma:string", "String");
        chunk.emit_op_u16(Op::CALL_IMPORT, to_str, line);
        chunk.emit(1, line);
        return;
    }

    // Two-arg path: [value, picture] → call __vybe_vb_format.
    // Trailing args beyond 2 are dropped (defensive — VB Format is
    // strictly 1-or-2-arity in practice).
    for _ in 2..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let global_name = chunk.add_constant(Value::String(Arc::from("__vybe_vb_format")));
    let value_slot = chunk.local_count;
    let picture_slot = value_slot + 1;
    chunk.local_count = value_slot + 2;
    // Stash picture (top), then value.
    chunk.emit_op_u16(Op::LOCAL_SET, picture_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunk.emit_op(Op::DROP, line);
    // Push global ref + (value, picture) and call.
    chunk.emit_op_u16(Op::GLOBAL_GET, global_name, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, picture_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 2, line);
}
