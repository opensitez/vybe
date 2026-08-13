//! `System.Windows.Forms.BindingSource` — the cursor over a data source.
//!
//! A `BindingSource` is not a control. It has no element, nothing paints it,
//! and every one of its members is data: a position, the list it walks, and the
//! four `Move*` verbs that step the position. It used to live in the WinForms
//! class table with `widget_host_fn: Some("new_BindingSource")`, which gave it a
//! `vybe:gui` backing object and routed every property through
//! `controlSetProperty`/`controlGetProperty` — a registry keyed by control name,
//! for a thing that is never a control. Reads answered from the object's own
//! fallback field when one had been written and from the GUI registry (a string,
//! or nothing) when one had not, so `bs.Position` read back `""` before anything
//! assigned it and `bs.MoveFirst()` was `undefined` — the class declared no
//! methods at all.
//!
//! So it is declared here, next to `DataTable`/`DataSet`/`DataRow`, as the data
//! object it is: a plain struct with real fields and real emits.
//!
//! ```text
//! { __type: "BindingSource", position, datasource, datamember, filter, sort }
//! ```
//!
//! Pattern: `emitter/core/datatable_adapter.rs`.

use std::sync::Arc;
use vybe_compiler::primitives::collections;
use vybe_compiler::primitives::ops;
use vybe_runtime::opcode::Op;
use vybe_runtime::opcode::heaptype::HT_EXTERN;
use vybe_runtime::{Chunk, Value};

const DATA_MEMBER_KEY: &str = "datamember";
const DATA_SOURCE_KEY: &str = "datasource";
const FILTER_KEY: &str = "filter";
const POSITION_KEY: &str = "position";
const ROWS_KEY: &str = "rows";
const SORT_KEY: &str = "sort";

/// Which end of the list a `Move*` verb aims at.
#[derive(Clone, Copy)]
pub enum Move {
    First,
    Next,
    Previous,
    Last,
}

fn reserve_slot(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn struct_get(chunk: &mut Chunk, object_local: u16, key: &str, line: u32) {
    let key_idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::LOCAL_GET, object_local, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key_idx, line);
}

fn struct_set_from_local(chunk: &mut Chunk, object_local: u16, key: &str, value: u16, line: u32) {
    let key_idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::LOCAL_GET, object_local, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key_idx, line);
}

/// `DUP → value → STRUCT_SET key`, leaving the object on the stack.
fn set_field(chunk: &mut Chunk, key: &str, val: impl FnOnce(&mut Chunk, u32), line: u32) {
    let key_idx = chunk.add_constant(Value::String(Arc::from(key)));
    vybe_compiler::primitives::instructions::core_wasm::dup(chunk, line);
    val(chunk, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key_idx, line);
}

/// `New BindingSource()` — every field the cursor owns, initialized.
///
/// `position` starts at `0` because .NET says so, and because the alternative
/// is what shipped: an unset field that read back as `""` and failed
/// `bs.Position = 0` on a freshly constructed source.
///
/// Stack in: `[]`   Stack out: `[obj]`
pub fn emit_bindingsource_new(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_struct_new(0, 0, line);
    set_field(
        chunk,
        "__type",
        |c, l| c.emit_string_const("BindingSource", l),
        line,
    );
    set_field(chunk, POSITION_KEY, |c, l| c.emit_f64_const(0.0, l), line);
    set_field(
        chunk,
        DATA_SOURCE_KEY,
        |c, l| c.emit_ref_null(HT_EXTERN, l),
        line,
    );
    set_field(
        chunk,
        DATA_MEMBER_KEY,
        |c, l| c.emit_string_const("", l),
        line,
    );
    set_field(chunk, FILTER_KEY, |c, l| c.emit_string_const("", l), line);
    set_field(chunk, SORT_KEY, |c, l| c.emit_string_const("", l), line);
}

/// The list the cursor walks, from whatever the data source is.
///
/// A `DataTable`/`DataSet` row container answers through its `rows` field; a
/// bare list IS the list; no source at all is an empty list. Written as two
/// nested value-`if`s rather than a single `STRUCT_GET`, because a data source
/// is routinely not a struct and `ecma:object.get` is the read that survives
/// that.
///
/// Stack in: `[]` (reads `bs_slot`)   Stack out: `[rows]`
fn emit_rows(chunks: &mut [Chunk], current: usize, bs_slot: u16, line: u32) {
    let (ds_slot, rows_slot) = {
        let chunk = &mut chunks[current];
        (reserve_slot(chunk), reserve_slot(chunk))
    };
    {
        let chunk = &mut chunks[current];
        struct_get(chunk, bs_slot, DATA_SOURCE_KEY, line);
        chunk.emit_op_u16(Op::LOCAL_SET, ds_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, ds_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_if_value(line);
    }
    collections::emit_array_new(chunks, current, 0, line);
    {
        let object_get = chunks[current].add_import("ecma:object", "get");
        let chunk = &mut chunks[current];
        chunk.emit_else(line);
        chunk.emit_op_u16(Op::LOCAL_GET, ds_slot, line);
        chunk.emit_string_const(ROWS_KEY, line);
        chunk.emit_call(object_get, 2, line);
        chunk.emit_op_u16(Op::LOCAL_SET, rows_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, rows_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_if_value(line);
        chunk.emit_op_u16(Op::LOCAL_GET, ds_slot, line);
        chunk.emit_else(line);
        chunk.emit_op_u16(Op::LOCAL_GET, rows_slot, line);
        chunk.emit_end(line);
        chunk.emit_end(line);
    }
}

/// `emit_rows` followed by a dynamic length, left in a fresh scratch slot.
fn emit_count_to_slot(chunks: &mut [Chunk], current: usize, bs_slot: u16, line: u32) -> u16 {
    emit_rows(chunks, current, bs_slot, line);
    collections::emit_len(chunks, current, line);
    let chunk = &mut chunks[current];
    let count_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line);
    count_slot
}

/// `bs.MoveFirst()` / `MoveNext()` / `MovePrevious()` / `MoveLast()`.
///
/// Every verb computes a candidate position and then clamps it into
/// `[0, count - 1]`, which is what makes all four safe on an empty source:
/// `count - 1` is `-1`, the low clamp pulls it back to `0`, and the position
/// never leaves the list.
///
/// Stack in: `[bs]`   Stack out: `[null]`
pub fn emit_bindingsource_move(chunks: &mut [Chunk], current: usize, mode: Move, line: u32) {
    let bs_slot = {
        let chunk = &mut chunks[current];
        let slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
        slot
    };
    let count_slot = emit_count_to_slot(chunks, current, bs_slot, line);
    let chunk = &mut chunks[current];
    let pos_slot = reserve_slot(chunk);

    match mode {
        Move::First => chunk.emit_f64_const(0.0, line),
        Move::Next => {
            struct_get(chunk, bs_slot, POSITION_KEY, line);
            chunk.emit_f64_const(1.0, line);
            ops::emit_dyn_add(chunk, line);
        }
        Move::Previous => {
            struct_get(chunk, bs_slot, POSITION_KEY, line);
            chunk.emit_f64_const(-1.0, line);
            ops::emit_dyn_add(chunk, line);
        }
        Move::Last => {
            chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
            chunk.emit_f64_const(-1.0, line);
            ops::emit_dyn_add(chunk, line);
        }
    }
    chunk.emit_op_u16(Op::LOCAL_SET, pos_slot, line);

    // Past the end → the last row.
    chunk.emit_op_u16(Op::LOCAL_GET, pos_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    ops::emit_dyn_ge(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunk.emit_f64_const(-1.0, line);
    ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, pos_slot, line);
    chunk.emit_end(line);

    // Before the start (including the empty-list `-1`) → the first row.
    chunk.emit_op_u16(Op::LOCAL_GET, pos_slot, line);
    chunk.emit_f64_const(0.0, line);
    ops::emit_dyn_lt(chunk, line);
    chunk.emit_if(line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, pos_slot, line);
    chunk.emit_end(line);

    struct_set_from_local(chunk, bs_slot, POSITION_KEY, pos_slot, line);
    chunk.emit_ref_null(HT_EXTERN, line);
}

/// `bs.Count` — how many rows the source has.
///
/// Stack in: `[bs]`   Stack out: `[count]`
pub fn emit_bindingsource_count(chunks: &mut [Chunk], current: usize, line: u32) {
    let bs_slot = {
        let chunk = &mut chunks[current];
        let slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
        slot
    };
    emit_rows(chunks, current, bs_slot, line);
    collections::emit_len(chunks, current, line);
}

/// `bs.Current` — the row at the cursor, or `null` when the source is empty.
///
/// The empty guard is not decoration: `ARRAY_GET` on an empty array traps, and
/// a data form asks for `Current` before it has anything to show.
///
/// Stack in: `[bs]`   Stack out: `[row | null]`
pub fn emit_bindingsource_current(chunks: &mut [Chunk], current: usize, line: u32) {
    let bs_slot = {
        let chunk = &mut chunks[current];
        let slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
        slot
    };
    emit_rows(chunks, current, bs_slot, line);
    let rows_slot = {
        let chunk = &mut chunks[current];
        let slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
        slot
    };
    collections::emit_len(chunks, current, line);
    let chunk = &mut chunks[current];
    let count_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunk.emit_i32_const(0, line);
    ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_ref_null(HT_EXTERN, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, rows_slot, line);
    struct_get(chunk, bs_slot, POSITION_KEY, line);
    chunk.emit_op(Op::I32_FROM_F64, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_end(line);
}
