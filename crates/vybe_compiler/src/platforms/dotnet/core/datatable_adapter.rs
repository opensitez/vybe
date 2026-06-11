//! .NET `System.Data` adapter — bytecode-only constructors.
//!
//! `DataTable`, `DataSet`, and `DataRow` constructors emit inline
//! bytecode (STRUCT_NEW + STRUCT_SET). No host fns added.
//! Methods on these types route through existing `vybe:data.*`
//! host fns via `DotnetClassExport` method bindings.
//!
//! Pattern: `emitter/dotnet/core/datetime_adapter.rs`.

use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

fn push_str(chunk: &mut Chunk, s: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(s)));
    chunk.emit_op_u16(Op::CONST, idx, line);
}

fn push_f64(chunk: &mut Chunk, v: f64, line: u32) {
    let idx = chunk.add_constant(Value::F64(v));
    chunk.emit_op_u16(Op::CONST, idx, line);
}

/// `DUP → push value → STRUCT_SET key → DROP`
/// Leaves the original object on the stack.
fn set_field(chunk: &mut Chunk, key: &str, val_fn: impl FnOnce(&mut Chunk, u32), line: u32) {
    let key_idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op(Op::DUP, line);
    val_fn(chunk, line);
    chunk.emit_op_u16(Op::STRUCT_SET, key_idx, line);
    chunk.emit_op(Op::DROP, line);
}

/// Reserve a local scratch slot.
fn reserve_slot(chunk: &mut Chunk) -> u16 {
    let slot = chunk.local_count;
    chunk.local_count = slot + 1;
    slot
}

/// `new DataTable(name?)` — creates `{ __type: "DataTable", tablename, columns: [], rows: [] }`.
///
/// Stack on entry: `[name]` (argc=1) or `[]` (argc=0 → name defaults to "Table1")
/// Stack on exit:  `[obj]`
pub fn emit_datatable_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];

    if argc == 0 {
        push_str(chunk, "Table1", line);
    }
    let name_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);

    set_field(chunk, "__type", |c, l| push_str(c, "DataTable", l), line);
    set_field(
        chunk,
        "tablename",
        |c, l| c.emit_op_u16(Op::LOCAL_GET, name_slot, l),
        line,
    );
    set_field(
        chunk,
        "columns",
        |c, l| c.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, l),
        line,
    );
    set_field(
        chunk,
        "rows",
        |c, l| c.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, l),
        line,
    );
    set_field(chunk, "count", |c, l| push_f64(c, 0.0, l), line);
}

/// `new DataSet(name?)` — creates `{ __type: "DataSet", datasetname, tables: [] }`.
///
/// Stack on entry: `[name]` (argc=1) or `[]` (argc=0 → name defaults to "DataSet1")
/// Stack on exit:  `[obj]`
pub fn emit_dataset_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];

    if argc == 0 {
        push_str(chunk, "DataSet1", line);
    }
    let name_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);

    set_field(chunk, "__type", |c, l| push_str(c, "DataSet", l), line);
    set_field(
        chunk,
        "datasetname",
        |c, l| c.emit_op_u16(Op::LOCAL_GET, name_slot, l),
        line,
    );
    set_field(
        chunk,
        "tables",
        |c, l| c.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, l),
        line,
    );
}

/// `new DataRow()` — creates `{ __type: "DataRow" }`.
///
/// Stack on entry: `[]` (no args)
/// Stack on exit:  `[obj]`
pub fn emit_datarow_new(chunk: &mut Chunk, line: u32) {
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    set_field(chunk, "__type", |c, l| push_str(c, "DataRow", l), line);
}
