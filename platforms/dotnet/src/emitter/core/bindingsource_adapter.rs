//! `System.Windows.Forms.BindingSource` — the cursor over a data source.
//!
//! A `BindingSource` is not a control. It has no element, nothing paints it,
//! and every one of its members is data: a position, the list it walks, and the
//! four `Move*` verbs that step the position. It used to live in the WinForms
//! class table with `widget_host_fn: Some("new_BindingSource")`, which gave it a
//! host backing object and routed every property through a generic accessor —
//! a registry keyed by control name,
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

use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};
use vybe_compiler::primitives::collections;
use vybe_compiler::primitives::ops;
use vybe_runtime::opcode::Op;
use vybe_runtime::opcode::heaptype::HT_EXTERN;
use vybe_runtime::{Chunk};

use super::object_fields::field_slot;

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
    chunk.emit_op_u16(Op::LOCAL_GET, object_local, line);
    class_slots::emit_class_get(chunk, ObjSource::Stack, &field_slot(key), Dest::Stack, line);
}

pub fn emit_get_field(chunk: &mut Chunk, key: &str, line: u32) {
    class_slots::emit_class_get(chunk, ObjSource::Stack, &field_slot(key), Dest::Stack, line);
}

fn struct_set_from_local(chunk: &mut Chunk, object_local: u16, key: &str, value: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, object_local, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(key),
        ValueSource::Stack,
        line,
    );
}

/// `DUP → value → STRUCT_SET key`, leaving the object on the stack.
fn set_field(chunk: &mut Chunk, key: &str, val: impl FnOnce(&mut Chunk, u32), line: u32) {
    vybe_compiler::primitives::instructions::core_wasm::dup(chunk, line);
    val(chunk, line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(key),
        ValueSource::Stack,
        line,
    );
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
    class_slots::emit_class_construct(
        chunk,
        "BindingSource",
        &[(field_slot(POSITION_KEY), ValueSource::ConstF64(0.0))],
        line,
    );
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
    set_field(
        chunk,
        "__bindings",
        |c, l| c.emit_array_new_fixed(0, 0, l),
        line,
    );
    set_field(
        chunk,
        "__grids",
        |c, l| c.emit_array_new_fixed(0, 0, l),
        line,
    );
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
        chunk.emit_op_u16(Op::LOCAL_GET, ds_slot, line);
        let undefined = chunk.add_import("wasm:js-undefined", "test");
        chunk.emit_call(undefined, 1, line);
        chunk.emit_op(Op::I32_OR, line);
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
        // A bare list has no rows property; its missing read is Undefined.
        chunk.emit_op_u16(Op::LOCAL_GET, rows_slot, line);
        let undefined = chunk.add_import("wasm:js-undefined", "test");
        chunk.emit_call(undefined, 1, line);
        chunk.emit_op(Op::I32_OR, line);
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
    emit_move(chunks, current, MoveInput::Static(mode), line);
}

/// Only the compiler-private navigator body passes this evaluated enum slot.
/// Public Move APIs and their exact inline emission remain unchanged.
pub(super) fn emit_bindingsource_move_from_private_mode(
    chunks: &mut [Chunk], current: usize, mode_slot: u16, line: u32,
) {
    emit_move(chunks, current, MoveInput::PrivateLocal(mode_slot), line);
}

enum MoveInput { Static(Move), PrivateLocal(u16) }

fn emit_position(chunk: &mut Chunk, bs_slot: u16, count_slot: u16, mode: Move, line: u32) {
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
}

fn emit_move(chunks: &mut [Chunk], current: usize, mode: MoveInput, line: u32) {
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
        MoveInput::Static(mode) => emit_position(chunk, bs_slot, count_slot, mode, line),
        MoveInput::PrivateLocal(slot) => {
            // Selector codes originate only in four generated wrappers.
            // These comparisons are raw standard i32 operations, not writable
            // dynamic comparison helpers; no user callback is added.
            for (code, mode) in [(0, Move::First), (1, Move::Next), (2, Move::Previous)] {
                chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
                chunk.emit_i32_const(code, line);
                chunk.emit_op(Op::I32_EQ, line);
                chunk.emit_if_value(line);
                emit_position(chunk, bs_slot, count_slot, mode, line);
                chunk.emit_else(line);
            }
            emit_position(chunk, bs_slot, count_slot, Move::Last, line);
            for _ in 0..3 { chunk.emit_end(line); }
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
    super::control_binding_adapter::refresh_controls(chunks, current, bs_slot, line);
    super::datagrid_binding_adapter::refresh_source_selection(chunks, current, bs_slot, line);
    super::bindingnavigator_adapter::refresh_source(chunks, current, bs_slot, line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
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

/// Stack: [source, data] -> [null].
pub fn emit_bindingsource_set_data_source(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let data = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, data, line);
    let source = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, source, line);
    struct_set_from_local(chunk, source, DATA_SOURCE_KEY, data, line);
    chunk.emit_op_u16(Op::LOCAL_GET, source, line);
    chunk.emit_f64_const(0.0, line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(POSITION_KEY),
        ValueSource::Stack,
        line,
    );
    super::control_binding_adapter::refresh(chunks, current, source, line);
    super::bindingnavigator_adapter::refresh_source(chunks, current, source, line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
}

/// Stack: [source, position] -> [null].
pub fn emit_bindingsource_set_position(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let position = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, position, line);
    let source = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, source, line);
    let count = emit_count_to_slot(chunks, current, source, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, position, line);
    chunk.emit_f64_const(0.0, line);
    ops::emit_dyn_lt(chunk, line);
    chunk.emit_if(line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, position, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, position, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count, line);
    ops::emit_dyn_ge(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, count, line);
    chunk.emit_f64_const(-1.0, line);
    ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, position, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, position, line);
    chunk.emit_f64_const(0.0, line);
    ops::emit_dyn_lt(chunk, line);
    chunk.emit_if(line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, position, line);
    chunk.emit_end(line);
    struct_set_from_local(chunk, source, POSITION_KEY, position, line);
    super::control_binding_adapter::refresh_controls(chunks, current, source, line);
    super::datagrid_binding_adapter::refresh_source_selection(chunks, current, source, line);
    super::bindingnavigator_adapter::refresh_source(chunks, current, source, line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
}

/// Stack: [source] -> [null].
pub fn emit_bindingsource_reset_bindings(chunks: &mut [Chunk], current: usize, line: u32) {
    let source = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, source, line);
    super::control_binding_adapter::refresh(chunks, current, source, line);
    super::bindingnavigator_adapter::refresh_source(chunks, current, source, line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
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
