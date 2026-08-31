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

use std::sync::Arc;
use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};
use vybe_compiler::primitives::instructions::{core_wasm, host};
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use vybe_compiler::primitives::math;

use super::object_fields::field_slot;

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        _ => panic!("push_const: no WASM-compliant encoding for {:?}", val),
    }
}

fn string_key(chunk: &mut Chunk, key: &str) -> u16 {
    if let Some((idx, _)) = chunk
        .constants
        .iter()
        .enumerate()
        .find(|(_, value)| matches!(value, Value::String(s) if s.as_ref() == key))
    {
        idx as u16
    } else {
        chunk.add_constant(Value::String(Arc::from(key)))
    }
}

fn struct_set_field(chunk: &mut Chunk, key: &str, line: u32) {
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(key),
        ValueSource::Stack,
        line,
    );
}

fn struct_set_named_field(chunk: &mut Chunk, key: &str, line: u32) {
    let value_slot = chunk.alloc_scratch(2);
    let obj_slot = value_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    push_const(chunk, Value::String(Arc::from(key)), line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    let idx = chunk.add_import("ecma:object", "set");
    // pushed a result it should not have; the VM was fixed and the other
    chunk.emit_call(idx, 3, line);
    // ECMA-262 §10.1.9 OrdinarySet RETURNS A BOOLEAN, so this call leaves a
    // value and the assignment's own result is `V` (§13.15.2), not it.
    // Removing this `DROP` made `++o.x` evaluate to null.
    chunk.emit_op(Op::DROP, line);
}

fn emit_array_get_const_index(chunk: &mut Chunk, array_slot: u16, index: f64, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, array_slot, line);
    push_const(chunk, Value::F64(index), line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

fn emit_parse_number_from_slot(chunks: &mut [Chunk], current: usize, text_slot: u16, line: u32) {
    let parse_int_idx = chunks[current].add_import("ecma:number", "parseInt");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunk.emit_call(parse_int_idx, 1, line);
    chunk.emit_op(Op::F64_FLOOR, line);
}

fn emit_store_array_part_as_number(
    chunks: &mut [Chunk],
    current: usize,
    array_slot: u16,
    index: f64,
    out_slot: u16,
    line: u32,
) {
    let chunk = &mut chunks[current];
    emit_array_get_const_index(chunk, array_slot, index, line);
    let text_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);
    emit_parse_number_from_slot(chunks, current, text_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
}

fn emit_total_ms_from_obj(chunk: &mut Chunk, obj_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    // ⛔ lowercase — `emit_build_timespan_from_total_ms` (the ONLY builder)
    // writes `totalmilliseconds`/`totalseconds`/`ticks`, so reading the
    // PascalCase spelling looked up a key nothing ever wrote. Write and read
    // must be one spelling; lowercase is the dotnet value-type convention.
    push_const(chunk, Value::String(Arc::from("totalmilliseconds")), line);
    let idx = chunk.add_import("ecma:object", "get");
    chunk.emit_call(idx, 2, line);
}

/// Build the TimeSpan object given the total milliseconds on the stack.
/// Stack on entry: `[total_ms]` ; Stack on exit: `[ts_obj]`
pub(crate) fn emit_build_timespan_from_total_ms(chunk: &mut Chunk, line: u32) {
    let ms_slot = chunk.alloc_scratch(6);
    let days_slot = ms_slot + 1;
    let rem_slot = ms_slot + 2;
    let hours_slot = ms_slot + 3;
    let minutes_slot = ms_slot + 4;
    let seconds_slot = ms_slot + 5;
    chunk.emit_op_u16(Op::LOCAL_SET, ms_slot, line);

    let type_key = string_key(chunk, "__type");
    let total_ms_key = string_key(chunk, "totalmilliseconds");
    let total_sec_key = string_key(chunk, "totalseconds");
    let total_min_key = string_key(chunk, "totalminutes");
    let total_hr_key = string_key(chunk, "totalhours");
    let total_day_key = string_key(chunk, "totaldays");
    let ticks_key = string_key(chunk, "ticks");

    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(86_400_000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    math::emit_trunc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, days_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, days_slot, line);
    push_const(chunk, Value::F64(86_400_000.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_SET, rem_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, rem_slot, line);
    push_const(chunk, Value::F64(3_600_000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    math::emit_trunc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, hours_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, rem_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, hours_slot, line);
    push_const(chunk, Value::F64(3_600_000.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_SET, rem_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, rem_slot, line);
    push_const(chunk, Value::F64(60_000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    math::emit_trunc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, minutes_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, rem_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, minutes_slot, line);
    push_const(chunk, Value::F64(60_000.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_SET, rem_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, rem_slot, line);
    push_const(chunk, Value::F64(1000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    math::emit_trunc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, seconds_slot, line);

    // The MILLISECONDS component — what is left after whole seconds. `rem_slot`
    // is finished with by this point, so it carries the answer rather than
    // costing a seventh scratch slot.
    chunk.emit_op_u16(Op::LOCAL_GET, rem_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, seconds_slot, line);
    push_const(chunk, Value::F64(1000.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_SET, rem_slot, line);

    let object_new = chunk.add_import("ecma:object", "new");
    chunk.emit_call(object_new, 0, line);

    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("TimeSpan")), line);
    struct_set_field(chunk, "__type", line);

    // ⛔ The COMPONENT properties — `Days`/`Hours`/`Minutes`/`Seconds`/
    // `Milliseconds`. Every one of them was already COMPUTED above, into
    // `days_slot`/`hours_slot`/`minutes_slot`/`seconds_slot`, and then THROWN
    // AWAY: only the `total*` keys and `ticks` were ever stored, so `ts.Days`
    // read `undefined` while `ts.TotalHours` answered correctly. The VB walker's
    // `fold_timespan_member_field` computed these at compile time and hid it.
    //
    // Both spellings, like every other field here: the lowercase key is what a
    // case-insensitive frontend folds to, the PascalCase one is .NET's own.
    for (slot, lower, pascal) in [
        (days_slot, "days", "Days"),
        (hours_slot, "hours", "Hours"),
        (minutes_slot, "minutes", "Minutes"),
        (seconds_slot, "seconds", "Seconds"),
        (rem_slot, "milliseconds", "Milliseconds"),
    ] {
        let key = string_key(chunk, lower);
        core_wasm::dup(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
        struct_set_field(chunk, lower, line);

        core_wasm::dup(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
        struct_set_named_field(chunk, pascal, line);
    }

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    struct_set_field(chunk, "totalmilliseconds", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    struct_set_named_field(chunk, "TotalMilliseconds", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(10_000.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    struct_set_field(chunk, "ticks", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(10_000.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    struct_set_named_field(chunk, "Ticks", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(1000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    struct_set_field(chunk, "totalseconds", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(1000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    struct_set_named_field(chunk, "TotalSeconds", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(60_000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    struct_set_field(chunk, "totalminutes", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(60_000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    struct_set_named_field(chunk, "TotalMinutes", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(3_600_000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    struct_set_field(chunk, "totalhours", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(3_600_000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    struct_set_named_field(chunk, "TotalHours", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(86_400_000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    struct_set_field(chunk, "totaldays", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(86_400_000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    struct_set_named_field(chunk, "TotalDays", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, days_slot, line);
    struct_set_named_field(chunk, "Days", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, hours_slot, line);
    struct_set_named_field(chunk, "Hours", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, minutes_slot, line);
    struct_set_named_field(chunk, "Minutes", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, seconds_slot, line);
    struct_set_named_field(chunk, "Seconds", line);

    // ⛔ `rem_slot` is ALREADY the remainder after whole seconds. The
    // component loop above is the only writer of `Milliseconds`; a second one
    // that subtracts the seconds again is wrong, and wrong only in the
    // PascalCase key, so a case-insensitive frontend cannot see it.
}

/// Build a TimeSpan from a count of `unit_ms` units. Stack: `[n]` →
/// `[ts]`. Internally: `total_ms = n * unit_ms`, then build the
/// record. Generic over unit so all `From*` methods share one body.
fn emit_timespan_from_unit(chunks: &mut Vec<Chunk>, current: usize, unit_ms: f64, line: u32) {
    let chunk = &mut chunks[current];
    push_const(chunk, Value::F64(unit_ms), line);
    chunk.emit_op(Op::F64_MUL, line);
    emit_build_timespan(chunks, current, line);
}

/// `TimeSpan.FromDays(n)` — `n * 86_400_000` ms.
pub fn emit_timespan_from_days(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_timespan_from_unit(chunks, current, 86_400_000.0, line);
}

/// `TimeSpan.FromHours(n)` — `n * 3_600_000` ms.
pub fn emit_timespan_from_hours(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_timespan_from_unit(chunks, current, 3_600_000.0, line);
}

/// `TimeSpan.FromMinutes(n)` — `n * 60_000` ms.
pub fn emit_timespan_from_minutes(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_timespan_from_unit(chunks, current, 60_000.0, line);
}

/// `TimeSpan.FromSeconds(n)` — `n * 1000` ms.
pub fn emit_timespan_from_seconds(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_timespan_from_unit(chunks, current, 1000.0, line);
}

/// `TimeSpan.FromMilliseconds(n)` — pass-through.
pub fn emit_timespan_from_milliseconds(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_timespan_from_unit(chunks, current, 1.0, line);
}

/// `TimeSpan.Zero` — 0-duration TimeSpan. Stack: `[]` → `[ts]`.
pub fn emit_timespan_zero(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    push_const(chunk, Value::F64(0.0), line);
    emit_build_timespan(chunks, current, line);
}

pub fn emit_timespan_parse(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let to_str_idx = chunks[current].add_import("ecma:string", "String");
    let chunk = &mut chunks[current];
    let text_slot = chunk.alloc_scratch(5);
    let parts_slot = text_slot + 1;
    let hours_slot = text_slot + 2;
    let minutes_slot = text_slot + 3;
    let seconds_slot = text_slot + 4;

    chunk.emit_call(to_str_idx, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    push_const(chunk, Value::String(Arc::from(":")), line);
    host::emit(chunk, "ecma:string", "split", 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, parts_slot, line);

    emit_store_array_part_as_number(chunks, current, parts_slot, 0.0, hours_slot, line);
    emit_store_array_part_as_number(chunks, current, parts_slot, 1.0, minutes_slot, line);
    emit_store_array_part_as_number(chunks, current, parts_slot, 2.0, seconds_slot, line);

    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, hours_slot, line);
    push_const(chunk, Value::F64(3600.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, minutes_slot, line);
    push_const(chunk, Value::F64(60.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_GET, seconds_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    push_const(chunk, Value::F64(1000.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    emit_build_timespan(chunks, current, line);
}

fn emit_compare_numeric_slots(chunk: &mut Chunk, left_slot: u16, right_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    chunk.emit_if(line);
    push_const(chunk, Value::I32(-1), line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    chunk.emit_if(line);
    push_const(chunk, Value::I32(1), line);
    chunk.emit_else(line);
    push_const(chunk, Value::I32(0), line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

pub fn emit_timespan_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    // .NET overloads by arity: (ticks) | (h,m,s) | (d,h,m,s) | (d,h,m,s,ms).
    // Everything reduces to total milliseconds; one builder serves all four.
    // The old body matched ONLY arity 3 and silently fell back to Zero for
    // the rest — `new TimeSpan(0,1,2,3)` compiled and answered 00:00:00.
    let chunk = &mut chunks[current];
    match argc {
        1 => {
            // Ticks: 100ns units, 10_000 per millisecond.
            push_const(chunk, Value::F64(10_000.0), line);
            chunk.emit_op(Op::F64_DIV, line);
            emit_build_timespan(chunks, current, line);
        }
        3 | 4 | 5 => {
            // Pop into slots right-to-left, then Horner up from the largest
            // unit present. Arity picks whether the leading arg is days.
            let n = argc as u16;
            let base = chunk.alloc_scratch(n);
            for i in 0..n {
                chunk.emit_op_u16(Op::LOCAL_SET, base + i, line);
            }
            // Slots now hold args reversed: base+0 = LAST arg.
            let has_days = argc >= 4;
            let has_ms = argc == 5;
            let mut next = n; // reading index, front to back
            let mut arg = |chunk: &mut Chunk| {
                next -= 1;
                chunk.emit_op_u16(Op::LOCAL_GET, base + next, line);
            };
            if has_days {
                arg(chunk);
                push_const(chunk, Value::F64(24.0), line);
                chunk.emit_op(Op::F64_MUL, line);
            } else {
                push_const(chunk, Value::F64(0.0), line);
            }
            // hours (+ days*24), then *60+minutes, *60+seconds, *1000(+ms)
            arg(chunk);
            chunk.emit_op(Op::F64_ADD, line);
            push_const(chunk, Value::F64(60.0), line);
            chunk.emit_op(Op::F64_MUL, line);
            arg(chunk);
            chunk.emit_op(Op::F64_ADD, line);
            push_const(chunk, Value::F64(60.0), line);
            chunk.emit_op(Op::F64_MUL, line);
            arg(chunk);
            chunk.emit_op(Op::F64_ADD, line);
            push_const(chunk, Value::F64(1000.0), line);
            chunk.emit_op(Op::F64_MUL, line);
            if has_ms {
                arg(chunk);
                chunk.emit_op(Op::F64_ADD, line);
            }
            emit_build_timespan(chunks, current, line);
        }
        _ => emit_timespan_zero(chunks, current, line),
    }
}

/// `TimeSpan.FromTicks(n)` — 100ns units.
pub fn emit_timespan_from_ticks(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    push_const(chunk, Value::F64(10_000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    emit_build_timespan(chunks, current, line);
}

/// `Int64::MAX` ticks in milliseconds — the magnitude of
/// `TimeSpan.MaxValue`/`MinValue` (±10,675,199 days). f64 carries it exactly
/// enough: the tests assert sign and scale, and .NET itself documents the
/// bound in ticks, not a full-precision ms value.
const MAX_VALUE_MS: f64 = 9.223372036854776e18 / 10_000.0;

pub fn emit_timespan_max_value(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    push_const(chunk, Value::F64(MAX_VALUE_MS), line);
    emit_build_timespan(chunks, current, line);
}

pub fn emit_timespan_min_value(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    push_const(chunk, Value::F64(-MAX_VALUE_MS), line);
    emit_build_timespan(chunks, current, line);
}

/// `ts.ToString()` — .NET's constant format: `[-][d.]hh:mm:ss[.fffffff]`.
/// Stack: `[ts]` → `[string]`.
pub fn emit_timespan_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj_slot = chunk.alloc_scratch(4);
    let ms_slot = obj_slot + 1;
    let part_slot = obj_slot + 2;
    let out_slot = obj_slot + 3;
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    emit_total_ms_from_obj(chunk, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ms_slot, line);

    // Sign prefix; work on |ms| after.
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_if(line);
    push_const(chunk, Value::String(Arc::from("-")), line);
    chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(-1.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ms_slot, line);
    chunk.emit_else(line);
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunk.emit_end(line);

    let number_to_string = chunk.add_import("ecma:number", "toString");
    let pad_start = chunk.add_import("ecma:string", "padStart");
    let concat = chunk.add_import("wasm:js-string", "concat");

    // days — printed WITHOUT padding, only when non-zero.
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(86_400_000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    math::emit_trunc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, part_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::F64_GT, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, part_slot, line);
    push_const(chunk, Value::I32(10), line);
    chunk.emit_call(number_to_string, 2, line);
    chunk.emit_call(concat, 2, line);
    push_const(chunk, Value::String(Arc::from(".")), line);
    chunk.emit_call(concat, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunk.emit_end(line);

    // hh:mm:ss — each: component, toString, padStart(2,'0'), concat.
    for (unit_ms, modulo, suffix) in [
        (3_600_000.0, 24.0, ":"),
        (60_000.0, 60.0, ":"),
        (1_000.0, 60.0, ""),
    ] {
        chunk.emit_op_u16(Op::LOCAL_GET, out_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
        push_const(chunk, Value::F64(unit_ms), line);
        chunk.emit_op(Op::F64_DIV, line);
        math::emit_trunc(chunk, line);
        push_const(chunk, Value::F64(modulo), line);
        math::emit_c_fmod(chunk, line);
        push_const(chunk, Value::I32(10), line);
        chunk.emit_call(number_to_string, 2, line);
        push_const(chunk, Value::I32(2), line);
        push_const(chunk, Value::String(Arc::from("0")), line);
        chunk.emit_call(pad_start, 3, line);
        chunk.emit_call(concat, 2, line);
        if !suffix.is_empty() {
            push_const(chunk, Value::String(Arc::from(suffix)), line);
            chunk.emit_call(concat, 2, line);
        }
        chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
    }

    // Fractional part — .NET prints seven digits (ticks precision), only
    // when the sub-second remainder is non-zero: ".fffffff" = ms*10_000.
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(1_000.0), line);
    math::emit_c_fmod(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, part_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::F64_GT, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, out_slot, line);
    push_const(chunk, Value::String(Arc::from(".")), line);
    chunk.emit_call(concat, 2, line);
    chunk.emit_op_u16(Op::LOCAL_GET, part_slot, line);
    push_const(chunk, Value::F64(10_000.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    math::emit_trunc(chunk, line);
    push_const(chunk, Value::I32(10), line);
    chunk.emit_call(number_to_string, 2, line);
    push_const(chunk, Value::I32(7), line);
    push_const(chunk, Value::String(Arc::from("0")), line);
    chunk.emit_call(pad_start, 3, line);
    chunk.emit_call(concat, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

pub fn emit_timespan_compare(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let right_slot = chunk.alloc_scratch(4);
    let left_slot = right_slot + 1;
    let right_ms_slot = right_slot + 2;
    let left_ms_slot = right_slot + 3;

    chunk.emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_slot, line);

    emit_total_ms_from_obj(chunk, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_ms_slot, line);
    emit_total_ms_from_obj(chunk, right_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, right_ms_slot, line);
    emit_compare_numeric_slots(chunk, left_ms_slot, right_ms_slot, line);
}

pub fn emit_timespan_negate(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    emit_total_ms_from_obj(chunk, obj_slot, line);
    push_const(chunk, Value::F64(-1.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    emit_build_timespan(chunks, current, line);
}

pub fn emit_timespan_duration(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    emit_total_ms_from_obj(chunk, obj_slot, line);
    math::emit_abs(chunk, line);
    emit_build_timespan(chunks, current, line);
}

pub fn emit_timespan_add(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let right_slot = chunk.alloc_scratch(2);
    let left_slot = right_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_slot, line);
    emit_total_ms_from_obj(chunk, left_slot, line);
    emit_total_ms_from_obj(chunk, right_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    emit_build_timespan(chunks, current, line);
}

pub fn emit_timespan_sub(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let right_slot = chunk.alloc_scratch(2);
    let left_slot = right_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_slot, line);
    emit_total_ms_from_obj(chunk, left_slot, line);
    emit_total_ms_from_obj(chunk, right_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    emit_build_timespan(chunks, current, line);
}

// ── Operator PROTOCOL SLOTS ──────────────────────────────────────────────────
//
// `ts + ts`, `ts - ts`, `ts * n`, `ts / n`, `ts / ts` and `-ts` are real .NET
// operators, and until now only ONE of them ever ran: the shared
// `try_compile_dotnet_datetime_timespan_binary_operator` catches `+`/`-` when
// BOTH operands' static types are known, so `a + b` on two typed locals worked
// while `a + TimeSpan.FromHours(2)` — the very same addition, with a CALL on
// the right — fell through to numeric addition and trapped in
// `wasm:js-number.toF64`. `*`, `/` and unary `-` had no arm at all.
//
// Answering it on the VALUE instead removes the question. `emit_rich_binop`
// (`primitives/operators.rs`) looks the ProtocolSlot up on the LEFT operand at
// run time and falls back to the primitive op when it is absent, so binding
// these makes every spelling reach the same body regardless of what the
// compiler could infer — which is what "a class should be ASKED, not matched
// by spelling" means for a value type.

fn create_function_chunk(name: &str, arity: u8) -> Chunk {
    let mut c = Chunk::new(name);
    c.arity = arity;
    c
}

/// The slots an arithmetic result must carry so the NEXT operator on it also
/// finds a body — `(a + b) + c`.
const TIMESPAN_OPERATOR_SLOTS: &[vybe_ast::ProtocolSlot] = &[
    vybe_ast::ProtocolSlot::Add,
    vybe_ast::ProtocolSlot::Sub,
    vybe_ast::ProtocolSlot::Mul,
    vybe_ast::ProtocolSlot::Div,
    vybe_ast::ProtocolSlot::Neg,
];

/// Copy the operator slots from `src_local` onto the object in `dst_local`.
///
/// ⛔ This is what keeps the binding non-recursive. A result built INSIDE an
/// operator body cannot bind fresh method chunks — that would need a new chunk
/// per operator per result, forever — so it inherits the receiver's, which are
/// the same functions.
fn emit_inherit_operator_slots(chunk: &mut Chunk, src_local: u16, dst_local: u16, line: u32) {
    for slot in TIMESPAN_OPERATOR_SLOTS {
        // A BOUND PROTOCOL SLOT, not a spelling. `ClassSlot::Slot` is the one
        // place a binding becomes a storage name, so these inherit under the
        // same identity the language bound them with.
        let bound = class_slots::resolve(
            &class_slots::ClassSlot::Slot(*slot),
            &class_slots::PlainNames,
        );
        class_slots::emit_class_get(chunk, ObjSource::Local(src_local), &bound, Dest::Stack, line);
        class_slots::emit_class_set(
            chunk,
            ObjSource::Local(dst_local),
            &bound,
            ValueSource::Stack,
            line,
        );
    }
}

/// Build a TimeSpan from total milliseconds on the stack INSIDE an operator
/// body, and give it the receiver's operator slots.
/// Stack: `[total_ms]` → `[ts]`.
fn emit_operator_result(chunk: &mut Chunk, receiver_local: u16, line: u32) {
    emit_build_timespan_from_total_ms(chunk, line);
    let result_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    emit_inherit_operator_slots(chunk, receiver_local, result_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// `a + b` / `a - b` on two TimeSpans.
fn push_timespan_addsub_chunk(chunks: &mut Vec<Chunk>, add: bool, line: u32) -> usize {
    let mut method = create_function_chunk(
        if add { "__timespan_add" } else { "__timespan_sub" },
        2,
    );
    method.local_count = 2;
    emit_total_ms_from_obj(&mut method, 0, line);
    emit_total_ms_from_obj(&mut method, 1, line);
    method.emit_op(if add { Op::F64_ADD } else { Op::F64_SUB }, line);
    emit_operator_result(&mut method, 0, line);
    method.emit_op(Op::RETURN, line);
    chunks.push(method);
    chunks.len() - 1
}

/// `a * n` — .NET's `TimeSpan * Double`. The right operand is a NUMBER; there
/// is no TimeSpan × TimeSpan.
fn push_timespan_mul_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut method = create_function_chunk("__timespan_mul", 2);
    method.local_count = 2;
    emit_total_ms_from_obj(&mut method, 0, line);
    method.emit_op_u16(Op::LOCAL_GET, 1, line);
    method.emit_op(Op::F64_MUL, line);
    emit_operator_result(&mut method, 0, line);
    method.emit_op(Op::RETURN, line);
    chunks.push(method);
    chunks.len() - 1
}

/// `a / b` — BOTH .NET overloads. `TimeSpan / Double` is a TimeSpan;
/// `TimeSpan / TimeSpan` is a `Double` ratio, so the divisor's kind decides
/// the RESULT TYPE and the branch is unavoidable.
fn push_timespan_div_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut method = create_function_chunk("__timespan_div", 2);
    method.local_count = 2;
    let result_slot = method.alloc_scratch(1);

    let typeof_fn = method.add_import("ecma:value", "typeof");
    method.emit_op_u16(Op::LOCAL_GET, 1, line);
    method.emit_call(typeof_fn, 1, line);
    method.emit_string_const("object", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut method, line);
    method.emit_if(line);
    // TimeSpan / TimeSpan → the ratio, a plain number.
    emit_total_ms_from_obj(&mut method, 0, line);
    emit_total_ms_from_obj(&mut method, 1, line);
    method.emit_op(Op::F64_DIV, line);
    method.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    method.emit_else(line);
    emit_total_ms_from_obj(&mut method, 0, line);
    method.emit_op_u16(Op::LOCAL_GET, 1, line);
    method.emit_op(Op::F64_DIV, line);
    emit_operator_result(&mut method, 0, line);
    method.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    method.emit_end(line);

    method.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    method.emit_op(Op::RETURN, line);
    chunks.push(method);
    chunks.len() - 1
}

/// `-a` — .NET `TimeSpan.Negate`.
fn push_timespan_neg_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut method = create_function_chunk("__timespan_neg", 1);
    method.local_count = 1;
    push_const(&mut method, Value::F64(0.0), line);
    emit_total_ms_from_obj(&mut method, 0, line);
    method.emit_op(Op::F64_SUB, line);
    emit_operator_result(&mut method, 0, line);
    method.emit_op(Op::RETURN, line);
    chunks.push(method);
    chunks.len() - 1
}

/// The chunk index of the operator body named `name`, creating it on first
/// use.
///
/// ⛔ ONE set of bodies per module, not one per construction site. These are
/// bound by EVERY `TimeSpan.From*` / `New TimeSpan` / `Zero` / `Parse` in the
/// program; minting five fresh chunks each time made the whole VB suite six
/// times slower to compile before this lookup went in. The chunks are pure
/// functions of their arguments, so sharing them is free.
fn timespan_operator_chunk(
    chunks: &mut Vec<Chunk>,
    name: &str,
    build: fn(&mut Vec<Chunk>, u32) -> usize,
    line: u32,
) -> usize {
    if let Some(idx) = chunks.iter().position(|c| c.name == name) {
        return idx;
    }
    build(chunks, line)
}

/// Bind `Add`/`Sub`/`Mul`/`Div`/`Neg` on the TimeSpan in `obj_slot`.
pub(crate) fn bind_timespan_operator_roles(
    chunks: &mut Vec<Chunk>,
    current: usize,
    obj_slot: u16,
    line: u32,
) {
    let bindings = [
        (
            vybe_ast::ProtocolSlot::Add,
            timespan_operator_chunk(
                chunks,
                "__timespan_add",
                |c, l| push_timespan_addsub_chunk(c, true, l),
                line,
            ),
        ),
        (
            vybe_ast::ProtocolSlot::Sub,
            timespan_operator_chunk(
                chunks,
                "__timespan_sub",
                |c, l| push_timespan_addsub_chunk(c, false, l),
                line,
            ),
        ),
        (
            vybe_ast::ProtocolSlot::Mul,
            timespan_operator_chunk(chunks, "__timespan_mul", push_timespan_mul_chunk, line),
        ),
        (
            vybe_ast::ProtocolSlot::Div,
            timespan_operator_chunk(chunks, "__timespan_div", push_timespan_div_chunk, line),
        ),
        (
            vybe_ast::ProtocolSlot::Neg,
            timespan_operator_chunk(chunks, "__timespan_neg", push_timespan_neg_chunk, line),
        ),
    ];
    for (slot, method_idx) in bindings {
        vybe_compiler::primitives::object::emit_bind_method(
            &mut chunks[current],
            obj_slot,
            &vybe_ast::protocol_slot_key(slot),
            method_idx,
            line,
        );
    }
}

/// Build a TimeSpan record from total milliseconds AND bind its operators.
/// Stack: `[total_ms]` → `[ts]`.
///
/// The entry point every caller that HAS the chunk list should use;
/// [`emit_build_timespan_from_total_ms`] alone leaves a value that cannot
/// answer `+`.
pub(crate) fn emit_build_timespan(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_build_timespan_from_total_ms(&mut chunks[current], line);
    let obj_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    bind_timespan_operator_roles(chunks, current, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}
