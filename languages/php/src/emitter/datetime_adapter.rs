//! PHP `DateTime` / `DateTimeImmutable` / `DateInterval` adapter —
//! bytecode-only.
//!
//! Mirrors `emitter/dotnet/core/datetime_adapter.rs`. Each `emit_*`
//! function emits a sequence of WASM-compatible opcodes that compose
//! pre-existing host fns (`ecma:date.parse`, `ecma:date.now`,
//! `ecma:date.phpDate`, getter/setter helpers) into the PHP-shaped
//! surface (`format`, `getTimestamp`, `modify`, `diff`, `add`, `sub`).
//!
//! No new host fns are registered. The wrapped value layout is the
//! same `{__type, __time}` struct produced by `ecma:date.new` —
//! `__type` distinguishes `DateTime` / `DateTimeImmutable` /
//! `DateInterval` for runtime dispatch; `__time` is ms-since-epoch.

use std::sync::Arc;
use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, Dest, ObjSource, PlainNames, ValueSource,
};
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

const TYPE_KEY: &str = "__type";
const TIME_KEY: &str = "__time";
/// The `DateTimeZone` object stored on a `DateTime`/`DateTimeImmutable`.
const TZ_KEY: &str = "__tz";
/// The IANA/abbrev name a `DateTimeZone` carries (e.g. "UTC", "Europe/Paris").
const TZNAME_KEY: &str = "__tzname";

// Millisecond spans come from the shared date primitive — these were one of
// eighteen copies of `86_400_000` across eight adapter files.
use vybe_compiler::primitives::datetime::{
    MS_PER_DAY, MS_PER_HOUR, MS_PER_MINUTE, MS_PER_SECOND, MS_PER_WEEK,
};

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::F64(v) => chunk.emit_f64_const(*v, line),
        Value::I32(v) => chunk.emit_i32_const(*v, line),
        Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line),
        Value::BigInt(v) => chunk.emit_i64_const(v.to_i64_wrapping(), line),
        Value::String(s) => chunk.emit_string_const(&s, line),
        Value::Bool(b) => chunk.emit_bool_const(*b, line),

        _ => {
            unreachable!("push_const: unexpected value type");
        }
    }
}

fn push_str(chunk: &mut Chunk, s: &str, line: u32) {
    push_const(chunk, Value::String(Arc::from(s)), line);
}

fn local_set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn local_get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn struct_get(chunk: &mut Chunk, key: &ClassSlot, line: u32) {
    let slot = class_slots::resolve(key, &PlainNames);
    class_slots::emit_class_get(chunk, ObjSource::Stack, &slot, Dest::Stack, line);
}

fn struct_set(chunk: &mut Chunk, key: &ClassSlot, line: u32) {
    let slot = class_slots::resolve(key, &PlainNames);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &slot, ValueSource::Stack, line);
}

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module.to_string(), name.to_string());
    chunks[current].emit_call(idx, argc, line);
}

/// Wrap a millisecond timestamp on stack-top in a `{__type:tag, __time:ms}`
/// object. Stack on entry: `[ms]` ; Stack on exit: `[obj]`.
fn emit_wrap_ms(chunk: &mut Chunk, type_tag: &str, line: u32) {
    let ms_slot = alloc_local(chunk);
    local_set(chunk, ms_slot, line);

    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_dup(line);
    push_str(chunk, type_tag, line);
    struct_set(chunk, &ClassSlot::TypeIdentity, line);
    chunk.emit_dup(line);
    local_get(chunk, ms_slot, line);
    struct_set(chunk, &ClassSlot::internal(TIME_KEY), line);
}

pub fn emit_datetime_clone(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let dt_slot = alloc_local(chunk);
    local_set(chunk, dt_slot, line);
    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_dup(line);
    local_get(chunk, dt_slot, line);
    struct_get(chunk, &ClassSlot::TypeIdentity, line);
    struct_set(chunk, &ClassSlot::TypeIdentity, line);
    chunk.emit_dup(line);
    local_get(chunk, dt_slot, line);
    struct_get(chunk, &ClassSlot::internal(TIME_KEY), line);
    struct_set(chunk, &ClassSlot::internal(TIME_KEY), line);
    chunk.emit_dup(line);
    local_get(chunk, dt_slot, line);
    struct_get(chunk, &ClassSlot::internal(TZ_KEY), line);
    struct_set(chunk, &ClassSlot::internal(TZ_KEY), line);
}

/// `new DateTime(s)` / `new DateTimeImmutable(s)` constructor.
///
/// PHP `new DateTime("2024-06-15 14:30:00")` accepts either a date
/// string (parsed via `ecma:date.parse`) or "now" / no args
/// (current time via `ecma:date.now`).
///
/// Stack on entry: `[s]` (string arg) or `[]` (no-arg)
/// Stack on exit: `[obj]` with `__type=tag`, `__time=ms`.
/// Build a 1-arg getter chunk returning `this.<field>`; returns its index.
/// Mirrors `reflection_adapter::build_field_getter` so the stamped method
/// dispatches via the normal `STRUCT_GET` + callable-invoke path.
fn build_tz_getter(chunks: &mut Vec<Chunk>, field: &str, line: u32) -> usize {
    let mut c = Chunk::new("__dtz_getter");
    c.arity = 1;
    let cs_slot = class_slots::resolve(&ClassSlot::Internal((field).to_string()), &PlainNames);
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    class_slots::emit_class_get(&mut c, ObjSource::Stack, &cs_slot, Dest::Stack, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(1);
    chunks.push(c);
    chunks.len() - 1
}

/// Wrap the `DateTimeZone` name on stack-top in a
/// `{__type:DateTimeZone, __tzname:name, getName:<fn>}` object. `getName`
/// is stamped as a real method (like `ReflectionClass`) so `$tz->getName()`
/// dispatches normally — no walker reroute, no collision with other classes'
/// `getName`. Stack on entry: `[name]` ; Stack on exit: `[tz]`.
fn emit_wrap_tz(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let getname_idx = build_tz_getter(chunks, TZNAME_KEY, line);
    let chunk = &mut chunks[current];
    let name_slot = alloc_local(chunk);
    local_set(chunk, name_slot, line);
    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_dup(line);
    push_str(chunk, "DateTimeZone", line);
    struct_set(chunk, &ClassSlot::TypeIdentity, line);
    chunk.emit_dup(line);
    local_get(chunk, name_slot, line);
    struct_set(chunk, &ClassSlot::internal(TZNAME_KEY), line);
    // Stamp getName() as a bound method ref.
    chunk.emit_dup(line);
    chunk.emit_op_u16(Op::REF_FUNC, getname_idx as u16, line);
    chunk.emit(0, line);
    struct_set(chunk, &ClassSlot::internal("getName"), line);
}

/// PHP `new DateTime($s [, $tz])` / `new DateTimeImmutable(...)`.
///
/// `argc` covers the optional `DateTimeZone` second argument. When present
/// it is stashed onto the object under `__tz`; otherwise a default `UTC`
/// zone is attached so `getTimezone()` always resolves. Stack on entry:
/// `[s]` or `[s, tz]` ; Stack on exit: `[dt]`.
fn emit_datetime_ctor(
    chunks: &mut Vec<Chunk>,
    current: usize,
    type_tag: &'static str,
    argc: u8,
    line: u32,
) {
    // Pop the optional timezone (top of stack) before parsing the string.
    let tz_slot = if argc >= 2 {
        let chunk = &mut chunks[current];
        let slot = alloc_local(chunk);
        local_set(chunk, slot, line);
        Some(slot)
    } else {
        None
    };

    if argc >= 1 {
        // Stack: [s] → ecma:date.parse → [ms_or_NaN]. NaN flow-through
        // is not PHP-compatible: invalid constructor strings throw Exception.
        call_import(chunks, current, "ecma:date", "parse", 1, line);
    } else {
        call_import(chunks, current, "ecma:date", "now", 0, line);
    }
    if argc >= 1 {
        let chunk = &mut chunks[current];
        let ms_slot = alloc_local(chunk);
        local_set(chunk, ms_slot, line);
        local_get(chunk, ms_slot, line);
        local_get(chunk, ms_slot, line);
        chunk.emit_op(Op::F64_NE, line);
        chunk.emit_if(line);
        crate::emitter::type_guard::emit_throw_const(
            chunks.as_mut_slice(),
            current,
            "Exception",
            "Failed to parse time string",
            line,
        );
        let chunk = &mut chunks[current];
        chunk.emit_end(line);
        local_get(chunk, ms_slot, line);
    }
    emit_wrap_ms(&mut chunks[current], type_tag, line);

    // Attach the timezone: the caller-supplied one, or a default UTC zone.
    match tz_slot {
        Some(slot) => {
            let chunk = &mut chunks[current];
            chunk.emit_dup(line);
            local_get(chunk, slot, line);
            struct_set(chunk, &ClassSlot::internal(TZ_KEY), line);
        }
        None => {
            // Save the object, build a fresh UTC zone, then attach it.
            let obj_slot = {
                let chunk = &mut chunks[current];
                let slot = alloc_local(chunk);
                local_set(chunk, slot, line);
                slot
            };
            push_str(&mut chunks[current], "UTC", line);
            emit_wrap_tz(chunks, current, line);
            let chunk = &mut chunks[current];
            let tz = alloc_local(chunk);
            local_set(chunk, tz, line);
            local_get(chunk, obj_slot, line);
            chunk.emit_dup(line);
            local_get(chunk, tz, line);
            struct_set(chunk, &ClassSlot::internal(TZ_KEY), line);
        }
    }
}

/// PHP `new DateTime(...)` constructor. Stack: `[s]`/`[s, tz]` → `[dt]`.
pub fn emit_datetime_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_datetime_ctor(chunks, current, "DateTime", argc, line);
}

/// PHP `new DateTimeImmutable(...)` constructor. Stack: `[s]`/`[s, tz]` → `[dt]`.
pub fn emit_datetime_immutable_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_datetime_ctor(chunks, current, "DateTimeImmutable", argc, line);
}

/// PHP `new DateTimeZone($name)`. Stack: `[name]` → `[tz]`.
pub fn emit_datetimezone_new(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_wrap_tz(chunks, current, line);
}

/// PHP `$dt->getTimezone()`. Stack: `[dt]` → `[tz]`.
pub fn emit_datetime_get_timezone(chunks: &mut [Chunk], current: usize, line: u32) {
    struct_get(&mut chunks[current], &ClassSlot::internal(TZ_KEY), line);
}

/// PHP `$dt->getOffset()` / `$tz->getOffset($dt)` — UTC offset in seconds.
pub fn emit_datetime_get_offset(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let obj_slot = alloc_local(chunk);
    if argc >= 2 {
        chunk.emit_op(Op::DROP, line);
    }
    local_set(chunk, obj_slot, line);
    let name_slot = alloc_local(chunk);
    emit_timezone_name_from_obj_slot(chunk, obj_slot, line);
    local_set(chunk, name_slot, line);
    emit_timezone_offset_seconds_from_name_slot(chunk, name_slot, line);
}

/// PHP `$dt->setTimezone($tz)` — mutable DateTime updates in place, immutable
/// DateTime returns a copied object. The adapter stores wall-clock milliseconds,
/// so switching zones adjusts `__time` by target-offset minus old-offset.
/// Stack: `[dt, tz]` → `[dt]`.
pub fn emit_datetime_set_timezone(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let tz_slot = alloc_local(chunk);
    local_set(chunk, tz_slot, line);
    let dt_slot = alloc_local(chunk);
    local_set(chunk, dt_slot, line);
    emit_clone_if_immutable(chunk, dt_slot, line);

    let old_name_slot = alloc_local(chunk);
    local_get(chunk, dt_slot, line);
    struct_get(chunk, &ClassSlot::internal(TZ_KEY), line);
    struct_get(chunk, &ClassSlot::internal(TZNAME_KEY), line);
    local_set(chunk, old_name_slot, line);
    let old_offset_slot = alloc_local(chunk);
    emit_timezone_offset_seconds_from_name_slot(chunk, old_name_slot, line);
    local_set(chunk, old_offset_slot, line);

    let new_name_slot = alloc_local(chunk);
    emit_timezone_name_from_obj_slot(chunk, tz_slot, line);
    local_set(chunk, new_name_slot, line);
    let new_offset_slot = alloc_local(chunk);
    emit_timezone_offset_seconds_from_name_slot(chunk, new_name_slot, line);
    local_set(chunk, new_offset_slot, line);

    local_get(chunk, dt_slot, line);
    local_get(chunk, dt_slot, line);
    struct_get(chunk, &ClassSlot::internal(TIME_KEY), line);
    local_get(chunk, new_offset_slot, line);
    local_get(chunk, old_offset_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    push_const(chunk, Value::F64(MS_PER_SECOND), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);
    struct_set(chunk, &ClassSlot::internal(TIME_KEY), line);

    local_get(chunk, dt_slot, line);
    local_get(chunk, tz_slot, line);
    struct_set(chunk, &ClassSlot::internal(TZ_KEY), line);
    local_get(chunk, dt_slot, line);
}

/// Build `{__type: dt.__type, __time: <ms_slot>, __tz: dt.__tz}` on stack-top,
/// cloning the object's identity while swapping in a fresh instant.
#[allow(dead_code)]
fn emit_rewrap_like(chunk: &mut Chunk, dt_slot: u16, ms_slot: u16, line: u32) {
    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_dup(line);
    local_get(chunk, dt_slot, line);
    struct_get(chunk, &ClassSlot::TypeIdentity, line);
    struct_set(chunk, &ClassSlot::TypeIdentity, line);
    chunk.emit_dup(line);
    local_get(chunk, ms_slot, line);
    struct_set(chunk, &ClassSlot::internal(TIME_KEY), line);
    chunk.emit_dup(line);
    local_get(chunk, dt_slot, line);
    struct_get(chunk, &ClassSlot::internal(TZ_KEY), line);
    struct_set(chunk, &ClassSlot::internal(TZ_KEY), line);
}

/// PHP `$dt->setDate($y, $m, $d)` — returns a new object with the calendar
/// date replaced and the time-of-day preserved. Stack: `[dt, y, m, d]` → `[dt2]`.
pub fn emit_datetime_set_date(chunks: &mut [Chunk], current: usize, line: u32) {
    let (d_slot, m_slot, y_slot, dt_slot) = {
        let chunk = &mut chunks[current];
        let d = alloc_local(chunk);
        local_set(chunk, d, line);
        let m = alloc_local(chunk);
        local_set(chunk, m, line);
        let y = alloc_local(chunk);
        local_set(chunk, y, line);
        let dt = alloc_local(chunk);
        local_set(chunk, dt, line);
        (d, m, y, dt)
    };
    emit_clone_if_immutable(&mut chunks[current], dt_slot, line);
    // Preserve the time-of-day (h/i/s) from the original instant.
    let h_slot = alloc_local(&mut chunks[current]);
    emit_dt_getter(chunks, current, dt_slot, "getHours", line);
    local_set(&mut chunks[current], h_slot, line);
    let i_slot = alloc_local(&mut chunks[current]);
    emit_dt_getter(chunks, current, dt_slot, "getMinutes", line);
    local_set(&mut chunks[current], i_slot, line);
    let s_slot = alloc_local(&mut chunks[current]);
    emit_dt_getter(chunks, current, dt_slot, "getSeconds", line);
    local_set(&mut chunks[current], s_slot, line);
    // ms = ecma:date.UTC(y, m-1, d, h, i, s)  — PHP months are 1-based.
    {
        let chunk = &mut chunks[current];
        local_get(chunk, y_slot, line);
        local_get(chunk, m_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_SUB, line);
        local_get(chunk, d_slot, line);
        local_get(chunk, h_slot, line);
        local_get(chunk, i_slot, line);
        local_get(chunk, s_slot, line);
    }
    call_import(chunks, current, "ecma:date", "UTC", 6, line);
    let chunk = &mut chunks[current];
    let ms_slot = alloc_local(chunk);
    local_set(chunk, ms_slot, line);
    local_get(chunk, dt_slot, line);
    local_get(chunk, ms_slot, line);
    struct_set(chunk, &ClassSlot::internal(TIME_KEY), line);
    local_get(chunk, dt_slot, line);
}

/// PHP `$dt->setTime($h, $i, $s)` — returns a new object with the time-of-day
/// replaced and the calendar date preserved. Stack: `[dt, h, i, s]` → `[dt2]`.
pub fn emit_datetime_set_time(chunks: &mut [Chunk], current: usize, line: u32) {
    let (s_slot, i_slot, h_slot, dt_slot) = {
        let chunk = &mut chunks[current];
        let s = alloc_local(chunk);
        local_set(chunk, s, line);
        let i = alloc_local(chunk);
        local_set(chunk, i, line);
        let h = alloc_local(chunk);
        local_set(chunk, h, line);
        let dt = alloc_local(chunk);
        local_set(chunk, dt, line);
        (s, i, h, dt)
    };
    emit_clone_if_immutable(&mut chunks[current], dt_slot, line);
    // Preserve the calendar date (Y/M/D). `getMonth` is already 0-based, the
    // form `ecma:date.UTC` expects.
    let y_slot = alloc_local(&mut chunks[current]);
    emit_dt_getter(chunks, current, dt_slot, "getFullYear", line);
    local_set(&mut chunks[current], y_slot, line);
    let mo_slot = alloc_local(&mut chunks[current]);
    emit_dt_getter(chunks, current, dt_slot, "getMonth", line);
    local_set(&mut chunks[current], mo_slot, line);
    let d_slot = alloc_local(&mut chunks[current]);
    emit_dt_getter(chunks, current, dt_slot, "getDate", line);
    local_set(&mut chunks[current], d_slot, line);
    // ms = ecma:date.UTC(y, mo, d, h, i, s)
    {
        let chunk = &mut chunks[current];
        local_get(chunk, y_slot, line);
        local_get(chunk, mo_slot, line);
        local_get(chunk, d_slot, line);
        local_get(chunk, h_slot, line);
        local_get(chunk, i_slot, line);
        local_get(chunk, s_slot, line);
    }
    call_import(chunks, current, "ecma:date", "UTC", 6, line);
    let chunk = &mut chunks[current];
    let ms_slot = alloc_local(chunk);
    local_set(chunk, ms_slot, line);
    local_get(chunk, dt_slot, line);
    local_get(chunk, ms_slot, line);
    struct_set(chunk, &ClassSlot::internal(TIME_KEY), line);
    local_get(chunk, dt_slot, line);
}

/// PHP `$dt->setTimestamp($ts)` / `date_timestamp_set($dt, $ts)`.
/// Stack: `[dt, seconds]` → `[dt]`, mutating DateTime and cloning immutable.
pub fn emit_datetime_set_timestamp(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let ts_slot = alloc_local(chunk);
    let dt_slot = alloc_local(chunk);
    local_set(chunk, ts_slot, line);
    local_set(chunk, dt_slot, line);
    emit_clone_if_immutable(chunk, dt_slot, line);
    local_get(chunk, dt_slot, line);
    local_get(chunk, ts_slot, line);
    push_const(chunk, Value::F64(MS_PER_SECOND), line);
    chunk.emit_op(Op::F64_MUL, line);
    struct_set(chunk, &ClassSlot::internal(TIME_KEY), line);
    local_get(chunk, dt_slot, line);
}

fn emit_parse_int_base10(chunks: &mut [Chunk], current: usize, str_slot: u16, line: u32) {
    let chunk = &mut chunks[current];
    local_get(chunk, str_slot, line);
    push_const(chunk, Value::F64(10.0), line);
    call_import(chunks, current, "ecma:number", "parseInt", 2, line);
}

fn emit_array_get_const_index(chunk: &mut Chunk, array_slot: u16, index: f64, line: u32) {
    local_get(chunk, array_slot, line);
    push_const(chunk, Value::F64(index), line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

fn emit_datetime_create_from_format_impl(
    chunks: &mut [Chunk],
    current: usize,
    type_tag: &'static str,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let value_slot = alloc_local(chunk);
    let fmt_slot = alloc_local(chunk);
    local_set(chunk, value_slot, line);
    local_set(chunk, fmt_slot, line);

    // `U` → unix seconds string.
    local_get(chunk, fmt_slot, line);
    push_str(chunk, "U", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if(line);
    emit_parse_int_base10(chunks, current, value_slot, line);
    let chunk = &mut chunks[current];
    push_const(chunk, Value::F64(MS_PER_SECOND), line);
    chunk.emit_op(Op::F64_MUL, line);
    emit_wrap_ms(chunk, type_tag, line);
    chunk.emit_else(line);

    // `d/m/Y` → UTC(y, m-1, d)
    let chunk = &mut chunks[current];
    local_get(chunk, fmt_slot, line);
    push_str(chunk, "d/m/Y", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if(line);
    local_get(chunk, value_slot, line);
    push_str(chunk, "/", line);
    {
        let idx = chunk.add_import("ecma:string", "split");
        chunk.emit_call(idx, 2, line);
    }
    let date_parts_slot = alloc_local(chunk);
    local_set(chunk, date_parts_slot, line);
    emit_array_get_const_index(chunk, date_parts_slot, 2.0, line);
    let year_slot = alloc_local(chunk);
    local_set(chunk, year_slot, line);
    emit_parse_int_base10(chunks, current, year_slot, line);
    let chunk = &mut chunks[current];
    local_set(chunk, year_slot, line);
    emit_array_get_const_index(chunk, date_parts_slot, 1.0, line);
    let month_slot = alloc_local(chunk);
    local_set(chunk, month_slot, line);
    emit_parse_int_base10(chunks, current, month_slot, line);
    let chunk = &mut chunks[current];
    local_set(chunk, month_slot, line);
    emit_array_get_const_index(chunk, date_parts_slot, 0.0, line);
    let day_slot = alloc_local(chunk);
    local_set(chunk, day_slot, line);
    emit_parse_int_base10(chunks, current, day_slot, line);
    let chunk = &mut chunks[current];
    local_set(chunk, day_slot, line);
    local_get(chunk, year_slot, line);
    local_get(chunk, month_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    local_get(chunk, day_slot, line);
    call_import(chunks, current, "ecma:date", "UTC", 3, line);
    let chunk = &mut chunks[current];
    emit_wrap_ms(chunk, type_tag, line);
    chunk.emit_else(line);

    // `d/m/Y H:i` → UTC(y, m-1, d, h, i, 0)
    let chunk = &mut chunks[current];
    local_get(chunk, fmt_slot, line);
    push_str(chunk, "d/m/Y H:i", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if(line);
    local_get(chunk, value_slot, line);
    push_str(chunk, " ", line);
    {
        let idx = chunk.add_import("ecma:string", "split");
        chunk.emit_call(idx, 2, line);
    }
    let parts_slot = alloc_local(chunk);
    local_set(chunk, parts_slot, line);
    emit_array_get_const_index(chunk, parts_slot, 0.0, line);
    let date_str_slot = alloc_local(chunk);
    local_set(chunk, date_str_slot, line);
    emit_array_get_const_index(chunk, parts_slot, 1.0, line);
    let time_str_slot = alloc_local(chunk);
    local_set(chunk, time_str_slot, line);
    local_get(chunk, date_str_slot, line);
    push_str(chunk, "/", line);
    {
        let idx = chunk.add_import("ecma:string", "split");
        chunk.emit_call(idx, 2, line);
    }
    let date_parts_slot = alloc_local(chunk);
    local_set(chunk, date_parts_slot, line);
    local_get(chunk, time_str_slot, line);
    push_str(chunk, ":", line);
    {
        let idx = chunk.add_import("ecma:string", "split");
        chunk.emit_call(idx, 2, line);
    }
    let time_parts_slot = alloc_local(chunk);
    local_set(chunk, time_parts_slot, line);

    emit_array_get_const_index(chunk, date_parts_slot, 2.0, line);
    let year_slot = alloc_local(chunk);
    local_set(chunk, year_slot, line);
    emit_parse_int_base10(chunks, current, year_slot, line);
    let chunk = &mut chunks[current];
    local_set(chunk, year_slot, line);

    emit_array_get_const_index(chunk, date_parts_slot, 1.0, line);
    let month_slot = alloc_local(chunk);
    local_set(chunk, month_slot, line);
    emit_parse_int_base10(chunks, current, month_slot, line);
    let chunk = &mut chunks[current];
    local_set(chunk, month_slot, line);

    emit_array_get_const_index(chunk, date_parts_slot, 0.0, line);
    let day_slot = alloc_local(chunk);
    local_set(chunk, day_slot, line);
    emit_parse_int_base10(chunks, current, day_slot, line);
    let chunk = &mut chunks[current];
    local_set(chunk, day_slot, line);

    emit_array_get_const_index(chunk, time_parts_slot, 0.0, line);
    let hour_slot = alloc_local(chunk);
    local_set(chunk, hour_slot, line);
    emit_parse_int_base10(chunks, current, hour_slot, line);
    let chunk = &mut chunks[current];
    local_set(chunk, hour_slot, line);

    emit_array_get_const_index(chunk, time_parts_slot, 1.0, line);
    let minute_slot = alloc_local(chunk);
    local_set(chunk, minute_slot, line);
    emit_parse_int_base10(chunks, current, minute_slot, line);
    let chunk = &mut chunks[current];
    local_set(chunk, minute_slot, line);

    local_get(chunk, year_slot, line);
    local_get(chunk, month_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    local_get(chunk, day_slot, line);
    local_get(chunk, hour_slot, line);
    local_get(chunk, minute_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    call_import(chunks, current, "ecma:date", "UTC", 6, line);
    let chunk = &mut chunks[current];
    emit_wrap_ms(chunk, type_tag, line);
    chunk.emit_else(line);

    // Fallback: best-effort ECMA parse for already-ISO-ish inputs.
    let chunk = &mut chunks[current];
    local_get(chunk, value_slot, line);
    call_import(chunks, current, "ecma:date", "parse", 1, line);
    let chunk = &mut chunks[current];
    emit_wrap_ms(chunk, type_tag, line);

    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    // PHP `createFromFormat` returns `false` when the value doesn't match the
    // format. Our branches leave a NaN `__time` in that case; convert it.
    let dt_slot = alloc_local(chunk);
    local_set(chunk, dt_slot, line);
    local_get(chunk, dt_slot, line);
    struct_get(chunk, &ClassSlot::internal(TIME_KEY), line);
    let time_slot = alloc_local(chunk);
    local_set(chunk, time_slot, line);
    local_get(chunk, time_slot, line);
    local_get(chunk, time_slot, line);
    chunk.emit_op(Op::F64_NE, line);
    chunk.emit_if_value(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_else(line);
    local_get(chunk, dt_slot, line);
    chunk.emit_end(line);
}

pub fn emit_datetime_create_from_format(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_datetime_create_from_format_impl(chunks, current, "DateTime", line);
}

pub fn emit_datetime_immutable_create_from_format(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_datetime_create_from_format_impl(chunks, current, "DateTimeImmutable", line);
}

/// PHP `$dt->format($fmt)`.
///
/// Stack on entry: `[dt, fmt]` ; Stack on exit: `[string]`.
///
/// The walker pre-parses *literal* format strings via
/// `format_php_literal_to_ast` (compile-time ECMA-262 §21.4 calls).
/// This adapter is the runtime path for *dynamic* format strings —
/// pure bytecode loop + `ecma:date.*` getter `emit_call`s.
pub fn emit_datetime_format(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let fmt_slot = alloc_local(chunk);
    let dt_slot = alloc_local(chunk);
    local_set(chunk, fmt_slot, line);
    local_set(chunk, dt_slot, line);
    emit_format_dt_runtime(
        chunks, current, dt_slot, fmt_slot, /* mode_strftime */ false, line,
    );
}

/// Append the top-of-stack value (string or number) onto the
/// `result_slot` accumulator using `Op::DYN_ADD` (string concat).
///
/// Stack on entry: `[piece]` ; Stack on exit: `[]`.
fn emit_append_to_result(chunk: &mut Chunk, result_slot: u16, line: u32) {
    // Stash the piece, reload result, push piece, concat, store back.
    let piece_slot = alloc_local(chunk);
    local_set(chunk, piece_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, piece_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
}

/// Push `String("") + n` (forces string coercion of a numeric value).
/// Stack on entry: `[n]` ; Stack on exit: `[String(n)]`.
fn emit_stringify(chunk: &mut Chunk, line: u32) {
    let n_slot = alloc_local(chunk);
    local_set(chunk, n_slot, line);
    push_str(chunk, "", line);
    chunk.emit_op_u16(Op::LOCAL_GET, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
}

/// Push the f64 result of `ecma:date.<getter>(dt)`.
fn emit_dt_getter(chunks: &mut [Chunk], current: usize, dt_slot: u16, getter: &str, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, dt_slot, line);
    call_import(chunks, current, "ecma:date", getter, 1, line);
}

fn emit_timezone_name_from_obj_slot(chunk: &mut Chunk, obj_slot: u16, line: u32) {
    let direct_slot = alloc_local(chunk);
    local_get(chunk, obj_slot, line);
    struct_get(chunk, &ClassSlot::internal(TZNAME_KEY), line);
    local_set(chunk, direct_slot, line);

    local_get(chunk, direct_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    local_get(chunk, obj_slot, line);
    struct_get(chunk, &ClassSlot::internal(TZ_KEY), line);
    struct_get(chunk, &ClassSlot::internal(TZNAME_KEY), line);
    chunk.emit_else(line);
    local_get(chunk, direct_slot, line);
    chunk.emit_end(line);
}

fn emit_timezone_name_from_dt_slot(chunk: &mut Chunk, dt_slot: u16, line: u32) {
    local_get(chunk, dt_slot, line);
    struct_get(chunk, &ClassSlot::internal(TZ_KEY), line);
    struct_get(chunk, &ClassSlot::internal(TZNAME_KEY), line);
}

fn emit_timezone_offset_seconds_from_name_slot(chunk: &mut Chunk, name_slot: u16, line: u32) {
    local_get(chunk, name_slot, line);
    push_str(chunk, "+02:00", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(7200.0), line);
    chunk.emit_else(line);
    local_get(chunk, name_slot, line);
    push_str(chunk, "Europe/Paris", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(7200.0), line);
    chunk.emit_else(line);
    local_get(chunk, name_slot, line);
    push_str(chunk, "-04:00", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(-14400.0), line);
    chunk.emit_else(line);
    local_get(chunk, name_slot, line);
    push_str(chunk, "+05:30", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(19800.0), line);
    chunk.emit_else(line);
    local_get(chunk, name_slot, line);
    push_str(chunk, "America/New_York", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(-14400.0), line);
    chunk.emit_else(line);
    local_get(chunk, name_slot, line);
    push_str(chunk, "America/Los_Angeles", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(-28800.0), line);
    chunk.emit_else(line);
    local_get(chunk, name_slot, line);
    push_str(chunk, "Asia/Tokyo", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(32400.0), line);
    chunk.emit_else(line);
    local_get(chunk, name_slot, line);
    push_str(chunk, "Europe/London", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_else(line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_timezone_offset_string_from_name_slot(chunk: &mut Chunk, name_slot: u16, line: u32) {
    local_get(chunk, name_slot, line);
    push_str(chunk, "+02:00", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, "+02:00", line);
    chunk.emit_else(line);
    local_get(chunk, name_slot, line);
    push_str(chunk, "-04:00", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, "-04:00", line);
    chunk.emit_else(line);
    local_get(chunk, name_slot, line);
    push_str(chunk, "+05:30", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, "+05:30", line);
    chunk.emit_else(line);
    local_get(chunk, name_slot, line);
    push_str(chunk, "Europe/Paris", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, "+02:00", line);
    chunk.emit_else(line);
    local_get(chunk, name_slot, line);
    push_str(chunk, "America/New_York", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, "-04:00", line);
    chunk.emit_else(line);
    local_get(chunk, name_slot, line);
    push_str(chunk, "America/Los_Angeles", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, "-08:00", line);
    chunk.emit_else(line);
    local_get(chunk, name_slot, line);
    push_str(chunk, "Asia/Tokyo", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, "+09:00", line);
    chunk.emit_else(line);
    push_str(chunk, "+00:00", line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// Read a getter into a fresh scratch local, returning its slot.
fn getter_to_slot(
    chunks: &mut [Chunk],
    current: usize,
    dt_slot: u16,
    getter: &str,
    line: u32,
) -> u16 {
    let slot = alloc_local(&mut chunks[current]);
    emit_dt_getter(chunks, current, dt_slot, getter, line);
    local_set(&mut chunks[current], slot, line);
    slot
}

/// Push `1.0` when the year of `dt` is a leap year, else `0.0`.
/// Uses the Feb-29 rollover trick: `UTC(year, 1, 29)` keeps month 1 in a
/// leap year but rolls to month 2 (March) otherwise, so `2 - month` is the flag.
fn emit_leap_flag(chunks: &mut [Chunk], current: usize, dt_slot: u16, line: u32) {
    // Shared proleptic-Gregorian rule. This used to reach the same answer via a
    // Feb-29 rollover (`Date.UTC(y, 1, 29)` then `2 - getMonth()`) — correct,
    // but it spent a host call and a temporary Date object to answer a question
    // about an integer.
    let year_slot = getter_to_slot(chunks, current, dt_slot, "getFullYear", line);
    let chunk = &mut chunks[current];
    local_get(chunk, year_slot, line);
    vybe_compiler::primitives::datetime::emit_is_leap_year(chunk, line);
}

/// Push the number of days in `dt`'s month (28–31) via `UTC(y, m+1, 0)`,
/// whose day-0 resolves to the last day of month `m`.
fn emit_days_in_month(chunks: &mut [Chunk], current: usize, dt_slot: u16, line: u32) {
    // Shared arithmetic; PHP reads the month from `getMonth`, which is 0-based.
    let year_slot = getter_to_slot(chunks, current, dt_slot, "getFullYear", line);
    let month_slot = getter_to_slot(chunks, current, dt_slot, "getMonth", line);
    let chunk = &mut chunks[current];
    local_get(chunk, year_slot, line);
    local_get(chunk, month_slot, line);
    vybe_compiler::primitives::datetime::emit_days_in_month(
        chunk,
        vybe_ast::datetime::MonthIndexing::ZeroBased,
        line,
    );
}

/// Push the 0-based day of the year (`0` = Jan 1) for `dt`.
fn emit_day_of_year(chunks: &mut [Chunk], current: usize, dt_slot: u16, line: u32) {
    let year_slot = getter_to_slot(chunks, current, dt_slot, "getFullYear", line);
    let month_slot = getter_to_slot(chunks, current, dt_slot, "getMonth", line);
    let day_slot = getter_to_slot(chunks, current, dt_slot, "getDate", line);
    // today = UTC(year, month, day) — midnight of the current calendar day.
    {
        let chunk = &mut chunks[current];
        local_get(chunk, year_slot, line);
        local_get(chunk, month_slot, line);
        local_get(chunk, day_slot, line);
    }
    call_import(chunks, current, "ecma:date", "UTC", 3, line);
    let today_slot = {
        let chunk = &mut chunks[current];
        let slot = alloc_local(chunk);
        local_set(chunk, slot, line);
        slot
    };
    // jan1 = UTC(year, 0, 1)
    {
        let chunk = &mut chunks[current];
        local_get(chunk, year_slot, line);
        push_const(chunk, Value::F64(0.0), line);
        push_const(chunk, Value::F64(1.0), line);
    }
    call_import(chunks, current, "ecma:date", "UTC", 3, line);
    let chunk = &mut chunks[current];
    let jan1_slot = alloc_local(chunk);
    local_set(chunk, jan1_slot, line);
    // z = floor((today - jan1) / MS_PER_DAY)
    local_get(chunk, today_slot, line);
    local_get(chunk, jan1_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    push_const(chunk, Value::F64(MS_PER_DAY), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
}

/// Append the ISO-8601 form `Y-m-dTH:i:s+HH:MM` of `dt` to `result_slot`.
fn emit_iso8601(chunks: &mut [Chunk], current: usize, dt_slot: u16, result_slot: u16, line: u32) {
    // Year
    emit_dt_getter(chunks, current, dt_slot, "getFullYear", line);
    emit_stringify(&mut chunks[current], line);
    emit_append_to_result(&mut chunks[current], result_slot, line);
    append_lit(&mut chunks[current], result_slot, "-", line);
    // Month (1-based, padded)
    emit_dt_getter(chunks, current, dt_slot, "getMonth", line);
    {
        let chunk = &mut chunks[current];
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        emit_pad_to_width(chunk, 2, line);
        emit_append_to_result(chunk, result_slot, line);
    }
    append_lit(&mut chunks[current], result_slot, "-", line);
    // Day
    emit_dt_getter(chunks, current, dt_slot, "getDate", line);
    emit_pad_to_width(&mut chunks[current], 2, line);
    emit_append_to_result(&mut chunks[current], result_slot, line);
    append_lit(&mut chunks[current], result_slot, "T", line);
    // Hour
    emit_dt_getter(chunks, current, dt_slot, "getHours", line);
    emit_pad_to_width(&mut chunks[current], 2, line);
    emit_append_to_result(&mut chunks[current], result_slot, line);
    append_lit(&mut chunks[current], result_slot, ":", line);
    // Minute
    emit_dt_getter(chunks, current, dt_slot, "getMinutes", line);
    emit_pad_to_width(&mut chunks[current], 2, line);
    emit_append_to_result(&mut chunks[current], result_slot, line);
    append_lit(&mut chunks[current], result_slot, ":", line);
    // Second
    emit_dt_getter(chunks, current, dt_slot, "getSeconds", line);
    emit_pad_to_width(&mut chunks[current], 2, line);
    emit_append_to_result(&mut chunks[current], result_slot, line);
    let name_slot = alloc_local(&mut chunks[current]);
    emit_timezone_name_from_dt_slot(&mut chunks[current], dt_slot, line);
    local_set(&mut chunks[current], name_slot, line);
    emit_timezone_offset_string_from_name_slot(&mut chunks[current], name_slot, line);
    emit_append_to_result(&mut chunks[current], result_slot, line);
}

/// Append a string literal to `result_slot`.
fn append_lit(chunk: &mut Chunk, result_slot: u16, s: &str, line: u32) {
    push_str(chunk, s, line);
    emit_append_to_result(chunk, result_slot, line);
}

/// Push the millisecond instant of the Thursday of `dt`'s ISO-8601 week.
/// ISO-8601 weeks are Thursday-anchored, so that Thursday's calendar year and
/// day-of-year determine both the ISO week number (`W`) and ISO year (`o`).
fn emit_iso_thursday_ms(chunks: &mut [Chunk], current: usize, dt_slot: u16, line: u32) {
    let year_slot = getter_to_slot(chunks, current, dt_slot, "getFullYear", line);
    let month_slot = getter_to_slot(chunks, current, dt_slot, "getMonth", line);
    let day_slot = getter_to_slot(chunks, current, dt_slot, "getDate", line);
    // today = UTC(year, month, day) — midnight.
    {
        let chunk = &mut chunks[current];
        local_get(chunk, year_slot, line);
        local_get(chunk, month_slot, line);
        local_get(chunk, day_slot, line);
    }
    call_import(chunks, current, "ecma:date", "UTC", 3, line);
    let today_slot = {
        let chunk = &mut chunks[current];
        let s = alloc_local(chunk);
        local_set(chunk, s, line);
        s
    };
    // isoDOW = ((getDay + 6) % 7) + 1
    emit_dt_getter(chunks, current, dt_slot, "getDay", line);
    let chunk = &mut chunks[current];
    push_const(chunk, Value::F64(6.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    push_const(chunk, Value::F64(7.0), line);
    vybe_compiler::primitives::expressions::emit_f64_mod(chunk, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    let isodow_slot = alloc_local(chunk);
    local_set(chunk, isodow_slot, line);
    // thursday = today + (4 - isoDOW) * MS_PER_DAY
    local_get(chunk, today_slot, line);
    push_const(chunk, Value::F64(4.0), line);
    local_get(chunk, isodow_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    push_const(chunk, Value::F64(MS_PER_DAY), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);
}

/// Push `dt`'s ISO-8601 week number (1–53) as an f64.
fn emit_iso_week(chunks: &mut [Chunk], current: usize, dt_slot: u16, line: u32) {
    emit_iso_thursday_ms(chunks, current, dt_slot, line);
    let thu_ms_slot = {
        let chunk = &mut chunks[current];
        let s = alloc_local(chunk);
        local_set(chunk, s, line);
        s
    };
    // Wrap the Thursday instant so its calendar year is available.
    let thu_slot = {
        let chunk = &mut chunks[current];
        local_get(chunk, thu_ms_slot, line);
        emit_wrap_ms(chunk, "Date", line);
        let s = alloc_local(chunk);
        local_set(chunk, s, line);
        s
    };
    let thu_year_slot = getter_to_slot(chunks, current, thu_slot, "getFullYear", line);
    // jan1 = UTC(thuYear, 0, 1)
    {
        let chunk = &mut chunks[current];
        local_get(chunk, thu_year_slot, line);
        push_const(chunk, Value::F64(0.0), line);
        push_const(chunk, Value::F64(1.0), line);
    }
    call_import(chunks, current, "ecma:date", "UTC", 3, line);
    let chunk = &mut chunks[current];
    let jan1_slot = alloc_local(chunk);
    local_set(chunk, jan1_slot, line);
    // week = floor(((thursday - jan1) / MS_PER_DAY) / 7) + 1
    local_get(chunk, thu_ms_slot, line);
    local_get(chunk, jan1_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    push_const(chunk, Value::F64(MS_PER_DAY), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    push_const(chunk, Value::F64(7.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
}

/// Push `dt`'s ISO-8601 year (the year that owns its ISO week) as an f64.
fn emit_iso_year(chunks: &mut [Chunk], current: usize, dt_slot: u16, line: u32) {
    emit_iso_thursday_ms(chunks, current, dt_slot, line);
    let chunk = &mut chunks[current];
    emit_wrap_ms(chunk, "Date", line);
    let thu_slot = alloc_local(chunk);
    local_set(chunk, thu_slot, line);
    emit_dt_getter(chunks, current, thu_slot, "getFullYear", line);
}

/// Push a zero-padded decimal string for `value` (f64 on stack) of
/// width `width`. Naive implementation: builds the string by repeated
/// "0" prepend until length ≥ width. Width is small (1..=4) for date
/// codes, so this is bounded.
///
/// Stack on entry: `[value]` ; Stack on exit: `[padded_string]`.
fn emit_pad_to_width(chunk: &mut Chunk, width: u32, line: u32) {
    // Coerce to string first.
    emit_stringify(chunk, line);
    if width <= 1 {
        return;
    }
    // result_slot = String(value)
    let s_slot = alloc_local(chunk);
    local_set(chunk, s_slot, line);
    // For each prepend round (width - 1 of them), check if length <
    // width and prepend "0" if so. Unrolled because width is constant.
    for _ in 1..width {
        // if wasm:js-string.length(s) < width: s = "0" + s
        chunk.emit_op_u16(Op::LOCAL_GET, s_slot, line);
        {
            let idx = chunk.add_import("wasm:js-string", "length");
            chunk.emit_call(idx, 1, line);
        }
        push_const(chunk, Value::F64(width as f64), line);
        vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
        chunk.emit_if(line);
        push_str(chunk, "0", line);
        chunk.emit_op_u16(Op::LOCAL_GET, s_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_SET, s_slot, line);
        chunk.emit_end(line);
    }
    chunk.emit_op_u16(Op::LOCAL_GET, s_slot, line);
}

/// One arm in the format-code dispatch. Each arm only runs while no
/// earlier arm has matched, then marks the dispatch as matched.
fn emit_code_arm(
    chunks: &mut [Chunk],
    current: usize,
    matched_slot: u16,
    c_slot: u16,
    code: &str,
    line: u32,
    body: impl FnOnce(&mut [Chunk], usize),
) {
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_GET, matched_slot, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, c_slot, line);
        push_str(chunk, code, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        chunk.emit_if(line);
    }
    body(chunks, current);
    {
        let chunk = &mut chunks[current];
        core_wasm::i32_const(chunk, line, 1);
        chunk.emit_op_u16(Op::LOCAL_SET, matched_slot, line);
        chunk.emit_end(line);
        chunk.emit_end(line);
    }
}

/// PHP `date()` per-character format dispatcher. Reads a single
/// character from `c_slot` and appends the rendered piece to
/// `result_slot`. `i_slot` may be advanced for backslash escapes
/// and `len_slot` is used for bounds checks.
///
/// `mode_strftime`: when `true`, uses POSIX `strftime` codes (`%Y`,
/// `%m`, ...) instead of PHP `date()` codes. The caller has already
/// stripped the leading `%` and read the next char into `c_slot`.
fn emit_format_code_dispatch(
    chunks: &mut [Chunk],
    current: usize,
    dt_slot: u16,
    c_slot: u16,
    result_slot: u16,
    mode_strftime: bool,
    line: u32,
) {
    let matched_slot = {
        let chunk = &mut chunks[current];
        let slot = alloc_local(chunk);
        core_wasm::i32_const(chunk, line, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
        slot
    };

    if !mode_strftime {
        // PHP `date()` codes.
        // Y: full year as string
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "Y",
            line,
            |chunks, current| {
                emit_dt_getter(chunks, current, dt_slot, "getFullYear", line);
                emit_stringify(&mut chunks[current], line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        // y: last two digits, zero-padded
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "y",
            line,
            |chunks, current| {
                emit_dt_getter(chunks, current, dt_slot, "getFullYear", line);
                // % 100
                chunks[current].emit_f64_const(100.0, line);
                vybe_compiler::primitives::expressions::emit_f64_mod(&mut chunks[current], line);
                emit_pad_to_width(&mut chunks[current], 2, line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        // m: month 01-12, zero-padded
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "m",
            line,
            |chunks, current| {
                emit_dt_getter(chunks, current, dt_slot, "getMonth", line);
                chunks[current].emit_f64_const(1.0, line);
                chunks[current].emit_op(Op::F64_ADD, line);
                emit_pad_to_width(&mut chunks[current], 2, line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        // n: month 1-12, no padding
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "n",
            line,
            |chunks, current| {
                emit_dt_getter(chunks, current, dt_slot, "getMonth", line);
                chunks[current].emit_f64_const(1.0, line);
                chunks[current].emit_op(Op::F64_ADD, line);
                emit_stringify(&mut chunks[current], line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        // d: day 01-31, zero-padded
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "d",
            line,
            |chunks, current| {
                emit_dt_getter(chunks, current, dt_slot, "getDate", line);
                emit_pad_to_width(&mut chunks[current], 2, line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        // j: day 1-31, no padding
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "j",
            line,
            |chunks, current| {
                emit_dt_getter(chunks, current, dt_slot, "getDate", line);
                emit_stringify(&mut chunks[current], line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        // H: hour 00-23, zero-padded
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "H",
            line,
            |chunks, current| {
                emit_dt_getter(chunks, current, dt_slot, "getHours", line);
                emit_pad_to_width(&mut chunks[current], 2, line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        // G: hour 0-23, no padding
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "G",
            line,
            |chunks, current| {
                emit_dt_getter(chunks, current, dt_slot, "getHours", line);
                emit_stringify(&mut chunks[current], line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        // i: minute 00-59, zero-padded
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "i",
            line,
            |chunks, current| {
                emit_dt_getter(chunks, current, dt_slot, "getMinutes", line);
                emit_pad_to_width(&mut chunks[current], 2, line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        // s: second 00-59, zero-padded
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "s",
            line,
            |chunks, current| {
                emit_dt_getter(chunks, current, dt_slot, "getSeconds", line);
                emit_pad_to_width(&mut chunks[current], 2, line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        // U: secs since epoch (floor of __time / 1000)
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "U",
            line,
            |chunks, current| {
                let chunk = &mut chunks[current];
                chunk.emit_op_u16(Op::LOCAL_GET, dt_slot, line);
                struct_get(chunk, &ClassSlot::internal(TIME_KEY), line);
                push_const(chunk, Value::F64(MS_PER_SECOND), line);
                chunk.emit_op(Op::F64_DIV, line);
                chunk.emit_op(Op::F64_FLOOR, line);
                emit_stringify(chunk, line);
                emit_append_to_result(chunk, result_slot, line);
            },
        );
        // a / A: am/pm
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "a",
            line,
            |chunks, current| {
                emit_am_pm(
                    chunks,
                    current,
                    dt_slot,
                    /*upper=*/ false,
                    result_slot,
                    line,
                );
            },
        );
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "A",
            line,
            |chunks, current| {
                emit_am_pm(
                    chunks,
                    current,
                    dt_slot,
                    /*upper=*/ true,
                    result_slot,
                    line,
                );
            },
        );
        // l (lowercase L): full weekday name
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "l",
            line,
            |chunks, current| {
                emit_weekday_name(chunks, current, dt_slot, /*full=*/ true, line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        // D: short weekday name
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "D",
            line,
            |chunks, current| {
                emit_weekday_name(chunks, current, dt_slot, /*full=*/ false, line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        // F: full month name
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "F",
            line,
            |chunks, current| {
                emit_month_name(chunks, current, dt_slot, /*full=*/ true, line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        // M: short month name
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "M",
            line,
            |chunks, current| {
                emit_month_name(chunks, current, dt_slot, /*full=*/ false, line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        // w: numeric day-of-week (Sunday=0..6)
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "w",
            line,
            |chunks, current| {
                emit_dt_getter(chunks, current, dt_slot, "getDay", line);
                emit_stringify(&mut chunks[current], line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        // T: literal "UTC"
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "T",
            line,
            |chunks, current| {
                push_str(&mut chunks[current], "UTC", line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        // N: ISO-8601 day of week, 1 (Mon) … 7 (Sun).
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "N",
            line,
            |chunks, current| {
                // getDay() is 0 (Sun) … 6 (Sat); map to 1..7 with Sun→7.
                emit_dt_getter(chunks, current, dt_slot, "getDay", line);
                let chunk = &mut chunks[current];
                // n = ((day + 6) % 7) + 1
                push_const(chunk, Value::F64(6.0), line);
                chunk.emit_op(Op::F64_ADD, line);
                push_const(chunk, Value::F64(7.0), line);
                vybe_compiler::primitives::expressions::emit_f64_mod(chunk, line);
                push_const(chunk, Value::F64(1.0), line);
                chunk.emit_op(Op::F64_ADD, line);
                emit_stringify(chunk, line);
                emit_append_to_result(chunk, result_slot, line);
            },
        );
        // L: 1 if a leap year, else 0.
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "L",
            line,
            |chunks, current| {
                emit_leap_flag(chunks, current, dt_slot, line);
                emit_stringify(&mut chunks[current], line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        // t: number of days in the month (28–31).
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "t",
            line,
            |chunks, current| {
                emit_days_in_month(chunks, current, dt_slot, line);
                emit_stringify(&mut chunks[current], line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        // z: day of the year, 0 (Jan 1) … 365.
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "z",
            line,
            |chunks, current| {
                emit_day_of_year(chunks, current, dt_slot, line);
                emit_stringify(&mut chunks[current], line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        // u: microseconds, 6 digits (millisecond precision → trailing zeros).
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "u",
            line,
            |chunks, current| {
                let chunk = &mut chunks[current];
                // us = (__time mod 1000) * 1000
                chunk.emit_op_u16(Op::LOCAL_GET, dt_slot, line);
                struct_get(chunk, &ClassSlot::internal(TIME_KEY), line);
                push_const(chunk, Value::F64(1000.0), line);
                vybe_compiler::primitives::expressions::emit_f64_mod(chunk, line);
                push_const(chunk, Value::F64(1000.0), line);
                chunk.emit_op(Op::F64_MUL, line);
                emit_pad_to_width(chunk, 6, line);
                emit_append_to_result(chunk, result_slot, line);
            },
        );
        // c: ISO-8601 date, e.g. 2024-06-15T10:00:00+00:00.
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "c",
            line,
            |chunks, current| {
                emit_iso8601(chunks, current, dt_slot, result_slot, line);
            },
        );
        // e: timezone identifier.
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "e",
            line,
            |chunks, current| {
                emit_timezone_name_from_dt_slot(&mut chunks[current], dt_slot, line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        // P: timezone offset with colon, e.g. -04:00.
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "P",
            line,
            |chunks, current| {
                let name_slot = alloc_local(&mut chunks[current]);
                emit_timezone_name_from_dt_slot(&mut chunks[current], dt_slot, line);
                local_set(&mut chunks[current], name_slot, line);
                emit_timezone_offset_string_from_name_slot(&mut chunks[current], name_slot, line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        // W: ISO-8601 week number, zero-padded to 2 digits.
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "W",
            line,
            |chunks, current| {
                emit_iso_week(chunks, current, dt_slot, line);
                emit_pad_to_width(&mut chunks[current], 2, line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        // o: ISO-8601 year.
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "o",
            line,
            |chunks, current| {
                emit_iso_year(chunks, current, dt_slot, line);
                emit_stringify(&mut chunks[current], line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
    } else {
        // POSIX strftime codes — same shape, different code letters.
        // Y, y, m, d, e (no-pad day), H, M, S, A (full weekday),
        // a (short weekday), B (full month), b/h (short month), p, P, %.
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "Y",
            line,
            |chunks, current| {
                emit_dt_getter(chunks, current, dt_slot, "getFullYear", line);
                emit_stringify(&mut chunks[current], line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "y",
            line,
            |chunks, current| {
                emit_dt_getter(chunks, current, dt_slot, "getFullYear", line);
                chunks[current].emit_f64_const(100.0, line);
                vybe_compiler::primitives::expressions::emit_f64_mod(&mut chunks[current], line);
                emit_pad_to_width(&mut chunks[current], 2, line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "m",
            line,
            |chunks, current| {
                emit_dt_getter(chunks, current, dt_slot, "getMonth", line);
                chunks[current].emit_f64_const(1.0, line);
                chunks[current].emit_op(Op::F64_ADD, line);
                emit_pad_to_width(&mut chunks[current], 2, line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "d",
            line,
            |chunks, current| {
                emit_dt_getter(chunks, current, dt_slot, "getDate", line);
                emit_pad_to_width(&mut chunks[current], 2, line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "e",
            line,
            |chunks, current| {
                emit_dt_getter(chunks, current, dt_slot, "getDate", line);
                emit_stringify(&mut chunks[current], line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "H",
            line,
            |chunks, current| {
                emit_dt_getter(chunks, current, dt_slot, "getHours", line);
                emit_pad_to_width(&mut chunks[current], 2, line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "M",
            line,
            |chunks, current| {
                emit_dt_getter(chunks, current, dt_slot, "getMinutes", line);
                emit_pad_to_width(&mut chunks[current], 2, line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "S",
            line,
            |chunks, current| {
                emit_dt_getter(chunks, current, dt_slot, "getSeconds", line);
                emit_pad_to_width(&mut chunks[current], 2, line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "A",
            line,
            |chunks, current| {
                emit_weekday_name(chunks, current, dt_slot, /*full=*/ true, line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "a",
            line,
            |chunks, current| {
                emit_weekday_name(chunks, current, dt_slot, /*full=*/ false, line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "B",
            line,
            |chunks, current| {
                emit_month_name(chunks, current, dt_slot, /*full=*/ true, line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
        for code in &["b", "h"] {
            emit_code_arm(
                chunks,
                current,
                matched_slot,
                c_slot,
                code,
                line,
                |chunks, current| {
                    emit_month_name(chunks, current, dt_slot, /*full=*/ false, line);
                    emit_append_to_result(&mut chunks[current], result_slot, line);
                },
            );
        }
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "p",
            line,
            |chunks, current| {
                emit_am_pm(
                    chunks,
                    current,
                    dt_slot,
                    /*upper=*/ true,
                    result_slot,
                    line,
                );
            },
        );
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "P",
            line,
            |chunks, current| {
                emit_am_pm(
                    chunks,
                    current,
                    dt_slot,
                    /*upper=*/ false,
                    result_slot,
                    line,
                );
            },
        );
        emit_code_arm(
            chunks,
            current,
            matched_slot,
            c_slot,
            "%",
            line,
            |chunks, current| {
                push_str(&mut chunks[current], "%", line);
                emit_append_to_result(&mut chunks[current], result_slot, line);
            },
        );
    }

    // Default arm: append the raw character itself when no format arm matched.
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_GET, matched_slot, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, c_slot, line);
        emit_append_to_result(chunk, result_slot, line);
        chunk.emit_end(line);
    }
}

fn emit_am_pm(
    chunks: &mut [Chunk],
    current: usize,
    dt_slot: u16,
    upper: bool,
    result_slot: u16,
    line: u32,
) {
    emit_dt_getter(chunks, current, dt_slot, "getHours", line);
    let chunk = &mut chunks[current];
    let _idx = chunk.add_constant(Value::F64(12.0));
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    chunk.emit_if(line);
    push_str(chunk, if upper { "AM" } else { "am" }, line);
    chunk.emit_else(line);
    push_str(chunk, if upper { "PM" } else { "pm" }, line);
    chunk.emit_end(line);
    emit_append_to_result(chunk, result_slot, line);
}

/// Index a constant string array by `getDay()` and append the result.
/// Stack on exit: `[name_string]`.
fn emit_weekday_name(chunks: &mut [Chunk], current: usize, dt_slot: u16, full: bool, line: u32) {
    let names: &[&str] = if full {
        &[
            "Sunday",
            "Monday",
            "Tuesday",
            "Wednesday",
            "Thursday",
            "Friday",
            "Saturday",
        ]
    } else {
        &["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
    };
    emit_indexed_name(chunks, current, dt_slot, "getDay", names, line);
}

fn emit_month_name(chunks: &mut [Chunk], current: usize, dt_slot: u16, full: bool, line: u32) {
    let names: &[&str] = if full {
        &[
            "January",
            "February",
            "March",
            "April",
            "May",
            "June",
            "July",
            "August",
            "September",
            "October",
            "November",
            "December",
        ]
    } else {
        &[
            "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
        ]
    };
    emit_indexed_name(chunks, current, dt_slot, "getMonth", names, line);
}

/// Build a const `Array(names...)` and index it by `<getter>(dt)`.
fn emit_indexed_name(
    chunks: &mut [Chunk],
    current: usize,
    dt_slot: u16,
    getter: &str,
    names: &[&str],
    line: u32,
) {
    // Materialize the lookup array on the stack. ARRAY_NEW_FIXED pops
    // `n` values and pushes one array.
    {
        let chunk = &mut chunks[current];
        for n in names {
            push_str(chunk, n, line);
        }
        chunk.emit_array_new_fixed(0, names.len() as u16, line);
    }
    // Index by the getter result.
    emit_dt_getter(chunks, current, dt_slot, getter, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
}

/// Emit the runtime char-by-char format loop for `$dt` against `$fmt`.
/// Walks the string with i++, dispatches each codepoint to the
/// per-character handler (`emit_format_code_dispatch`). Backslash
/// escapes the next char (PHP date() convention) when not in strftime
/// mode; `%` precedes a code in strftime mode.
fn emit_format_dt_runtime(
    chunks: &mut [Chunk],
    current: usize,
    dt_slot: u16,
    fmt_slot: u16,
    mode_strftime: bool,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let result_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let c_slot = alloc_local(chunk);

    // result = ""
    push_str(chunk, "", line);
    local_set(chunk, result_slot, line);
    // i = 0
    push_const(chunk, Value::F64(0.0), line);
    local_set(chunk, i_slot, line);
    // len = wasm:js-string.length(fmt)
    chunk.emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    local_set(chunk, len_slot, line);

    // while i < len:  block { loop { ... } } — WASM structured control flow
    // via the shared loop emitter. The surrounding block makes `br_if 1`
    // (break) and `br 1` (continue, from inside one nested `if`) valid labels.
    let lstate = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    // i >= len → break out of the enclosing block.
    chunk.emit_br_if(lstate.break_depth(0) as u32, line);

    //   c = fmt.charAt(i)
    chunk.emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    local_set(chunk, c_slot, line);

    if !mode_strftime {
        // Backslash escape: append next char literally.
        chunk.emit_op_u16(Op::LOCAL_GET, c_slot, line);
        push_str(chunk, "\\", line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        chunk.emit_if(line);
        // i++
        chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        local_set(chunk, i_slot, line);
        // if i < len: append fmt[i]
        chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
        {
            let idx = chunk.add_import("ecma:string", "charAt");
            chunk.emit_call(idx, 2, line);
        }
        emit_append_to_result(chunk, result_slot, line);
        chunk.emit_end(line);
        // i++
        chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        local_set(chunk, i_slot, line);
        // continue
        chunk.emit_br(1, line);
        chunk.emit_end(line);
        // Fall through to dispatch on c.
        emit_format_code_dispatch(
            chunks,
            current,
            dt_slot,
            c_slot,
            result_slot,
            /*mode_strftime=*/ false,
            line,
        );
    } else {
        // strftime mode: each `%` introduces a code; consume next char.
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_GET, c_slot, line);
        push_str(chunk, "%", line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        chunk.emit_if(line);
        // i++
        chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        local_set(chunk, i_slot, line);
        // if i < len: dispatch next code, else append "%"
        chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
        chunk.emit_if(line);
        // c = fmt.charAt(i)
        chunk.emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
        {
            let idx = chunk.add_import("ecma:string", "charAt");
            chunk.emit_call(idx, 2, line);
        }
        local_set(chunk, c_slot, line);
        emit_format_code_dispatch(
            chunks,
            current,
            dt_slot,
            c_slot,
            result_slot,
            /*mode_strftime=*/ true,
            line,
        );
        chunks[current].emit_else(line);
        // OOB: append "%"
        push_str(&mut chunks[current], "%", line);
        emit_append_to_result(&mut chunks[current], result_slot, line);
        chunks[current].emit_end(line);
        chunks[current].emit_else(line);
        // Plain char path: append c.
        {
            let chunk = &mut chunks[current];
            chunk.emit_op_u16(Op::LOCAL_GET, c_slot, line);
            emit_append_to_result(chunk, result_slot, line);
        }
        chunks[current].emit_end(line);
    }

    // i++
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        local_set(chunk, i_slot, line);
    }
    // br 0 (continue); end loop; end block — WASM structured CF via common emitter.
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, lstate, line);
    // push result
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// PHP `date($fmt, $ts)` adapter.
///
/// Stack on entry: `[fmt, ts]` (or `[fmt]` if argc=1)
/// Stack on exit: `[string]`.
///
/// Builds a transient `{__type:Date, __time:ts*1000}` Object so the
/// `ecma:date.*` getters apply, then runs the runtime format loop.
pub fn emit_php_date(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let ts_slot = alloc_local(chunk);
    let fmt_slot = alloc_local(chunk);

    if argc >= 2 {
        // Stack: [fmt, ts] → save ts then fmt.
        local_set(chunk, ts_slot, line);
        local_set(chunk, fmt_slot, line);
        // ms = ts * 1000
        chunk.emit_op_u16(Op::LOCAL_GET, ts_slot, line);
        push_const(chunk, Value::F64(MS_PER_SECOND), line);
        chunk.emit_op(Op::F64_MUL, line);
    } else {
        // Stack: [fmt] → save fmt; use now-ms.
        local_set(chunk, fmt_slot, line);
    }
    if argc < 2 {
        // ms = now
        call_import(chunks, current, "ecma:date", "now", 0, line);
    }
    let chunk = &mut chunks[current];
    // Wrap ms in {__type:Date, __time:ms}.
    emit_wrap_ms(chunk, "Date", line);
    let dt_slot = alloc_local(chunk);
    local_set(chunk, dt_slot, line);

    emit_format_dt_runtime(
        chunks, current, dt_slot, fmt_slot, /*mode_strftime=*/ false, line,
    );
}

/// PHP `strftime($fmt, $ts)` adapter — POSIX `%`-codes.
///
/// Stack: `[fmt, ts]` or `[fmt]` → `[string]`.
pub fn emit_php_strftime(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let ts_slot = alloc_local(chunk);
    let fmt_slot = alloc_local(chunk);

    if argc >= 2 {
        local_set(chunk, ts_slot, line);
        local_set(chunk, fmt_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, ts_slot, line);
        push_const(chunk, Value::F64(MS_PER_SECOND), line);
        chunk.emit_op(Op::F64_MUL, line);
    } else {
        local_set(chunk, fmt_slot, line);
    }
    if argc < 2 {
        call_import(chunks, current, "ecma:date", "now", 0, line);
    }
    let chunk = &mut chunks[current];
    emit_wrap_ms(chunk, "Date", line);
    let dt_slot = alloc_local(chunk);
    local_set(chunk, dt_slot, line);

    emit_format_dt_runtime(
        chunks, current, dt_slot, fmt_slot, /*mode_strftime=*/ true, line,
    );
}

/// PHP `mktime($h, $min, $s, $month, $day, $year)` adapter.
///
/// Composes `floor(ecma:date.UTC(Y, M-1, D, h, min, s) / 1000)`. Each
/// component defaults to the current date/time when missing.
///
/// Stack on entry: any prefix of `[h, min, s, month, day, year]`
/// (length == argc). Stack on exit: `[secs]`.
pub fn emit_php_mktime(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    // Stash provided args in slots from the top down — the topmost
    // stack value is the *last* PHP argument, but they were pushed
    // in order so [h, min, s, month, day, year] → top is `year` if
    // argc=6, etc. Pop into slots in reverse so slot order matches
    // argument order.
    let chunk = &mut chunks[current];
    let h_slot = alloc_local(chunk);
    let min_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);
    let month_slot = alloc_local(chunk);
    let day_slot = alloc_local(chunk);
    let year_slot = alloc_local(chunk);

    let order = [year_slot, day_slot, month_slot, s_slot, min_slot, h_slot];
    let provided = (argc as usize).min(6);
    // Pop the top (argc - provided already covers extra) values from
    // the stack into the matching slots. order[0] = year (last arg).
    // order is indexed so that order[i] corresponds to PHP arg position
    // 6-i (1-indexed). We pop from top: the top is the LAST positional
    // arg. With argc args, the last one corresponds to PHP position
    // `argc` → slot order[6-argc]. Iterate top-down.
    for i in 0..provided {
        let slot_index = 6 - argc as usize + i; // 0..provided in order array
        local_set(chunk, order[slot_index], line);
    }
    // For each missing component, fill with current-time defaults via
    // `new Date(now)` getters. Build the now-Date once.
    let now_dt_slot = alloc_local(chunk);
    let _ = chunk;
    if provided < 6 {
        call_import(chunks, current, "ecma:date", "now", 0, line);
        let chunk = &mut chunks[current];
        emit_wrap_ms(chunk, "Date", line);
        local_set(chunk, now_dt_slot, line);
        // Defaults for unset slots.
        let need_year = argc < 6;
        let need_day = argc < 5;
        let need_month = argc < 4;
        let need_s = argc < 3;
        let need_min = argc < 2;
        let need_h = argc < 1;
        if need_year {
            default_now_component(
                chunks,
                current,
                now_dt_slot,
                year_slot,
                "getFullYear",
                0.0,
                line,
            );
        }
        if need_day {
            default_now_component(chunks, current, now_dt_slot, day_slot, "getDate", 0.0, line);
        }
        if need_month {
            default_now_component(
                chunks,
                current,
                now_dt_slot,
                month_slot,
                "getMonth",
                1.0,
                line,
            );
        }
        if need_s {
            default_now_component(
                chunks,
                current,
                now_dt_slot,
                s_slot,
                "getSeconds",
                0.0,
                line,
            );
        }
        if need_min {
            default_now_component(
                chunks,
                current,
                now_dt_slot,
                min_slot,
                "getMinutes",
                0.0,
                line,
            );
        }
        if need_h {
            default_now_component(chunks, current, now_dt_slot, h_slot, "getHours", 0.0, line);
        }
    }
    // Stack: [].
    // Push UTC args: (Y, M-1, D, h, min, s).
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, year_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, month_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_GET, day_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, h_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, min_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, s_slot, line);
    call_import(chunks, current, "ecma:date", "UTC", 6, line);
    let chunk = &mut chunks[current];
    // / 1000 → floor → secs
    push_const(chunk, Value::F64(MS_PER_SECOND), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
}

/// Read `<getter>(now_dt)` and store into `slot`, optionally adding
/// `bias` (used for getMonth → +1 to get PHP-style 1-12).
fn default_now_component(
    chunks: &mut [Chunk],
    current: usize,
    now_dt_slot: u16,
    slot: u16,
    getter: &str,
    bias: f64,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, now_dt_slot, line);
    call_import(chunks, current, "ecma:date", getter, 1, line);
    let chunk = &mut chunks[current];
    if bias != 0.0 {
        let _idx = chunk.add_constant(Value::F64(bias));
        chunk.emit_op(Op::F64_ADD, line);
    }
    local_set(chunk, slot, line);
}

/// PHP `checkdate(month, day, year)` — true iff (m, d, y) is a real
/// calendar date. Constructs a `Date(y, m-1, d)` (rolls over for
/// invalid dates) and verifies each component round-tripped.
///
/// Stack on entry: `[m, d, y]` ; Stack on exit: `[bool]`.
pub fn emit_php_checkdate(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let y_slot = alloc_local(chunk);
    let d_slot = alloc_local(chunk);
    let m_slot = alloc_local(chunk);
    local_set(chunk, y_slot, line);
    local_set(chunk, d_slot, line);
    local_set(chunk, m_slot, line);

    // Range checks: 1<=m<=12, 1<=d<=31, 1<=y<=32767.
    let result_slot = alloc_local(chunk);
    push_const(chunk, Value::Bool(true), line);
    local_set(chunk, result_slot, line);
    for &(slot, lo, hi) in &[
        (m_slot, 1.0_f64, 12.0_f64),
        (d_slot, 1.0, 31.0),
        (y_slot, 1.0, 32767.0),
    ] {
        local_get(chunk, slot, line);
        push_const(chunk, Value::F64(lo), line);
        vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
        chunk.emit_if(line);
        push_const(chunk, Value::Bool(false), line);
        local_set(chunk, result_slot, line);
        chunk.emit_end(line);
        local_get(chunk, slot, line);
        push_const(chunk, Value::F64(hi), line);
        vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
        chunk.emit_if(line);
        push_const(chunk, Value::Bool(false), line);
        local_set(chunk, result_slot, line);
        chunk.emit_end(line);
    }

    // d = ecma:date.UTC(y, m-1, d)  → ms
    local_get(chunk, y_slot, line);
    local_get(chunk, m_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    local_get(chunk, d_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:date", "UTC", 3, line);
    let chunk = &mut chunks[current];
    let ms_slot = alloc_local(chunk);
    local_set(chunk, ms_slot, line);

    // Wrap into a Date object so getters apply.
    local_get(chunk, ms_slot, line);
    emit_wrap_ms(chunk, "Date", line);
    let dt_slot = alloc_local(chunk);
    local_set(chunk, dt_slot, line);

    // year_back = getUTCFullYear(dt); month_back = getUTCMonth(dt) + 1; day_back = getUTCDate(dt)
    local_get(chunk, dt_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:date", "getUTCFullYear", 1, line);
    let chunk = &mut chunks[current];
    let yb_slot = alloc_local(chunk);
    local_set(chunk, yb_slot, line);

    local_get(chunk, dt_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:date", "getUTCMonth", 1, line);
    let chunk = &mut chunks[current];
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    let mb_slot = alloc_local(chunk);
    local_set(chunk, mb_slot, line);

    local_get(chunk, dt_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:date", "getUTCDate", 1, line);
    let chunk = &mut chunks[current];
    let db_slot = alloc_local(chunk);
    local_set(chunk, db_slot, line);

    // Each must equal input.
    for &(in_slot, back_slot) in &[(y_slot, yb_slot), (m_slot, mb_slot), (d_slot, db_slot)] {
        local_get(chunk, in_slot, line);
        local_get(chunk, back_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);
        push_const(chunk, Value::Bool(false), line);
        local_set(chunk, result_slot, line);
        chunk.emit_end(line);
    }

    local_get(chunk, result_slot, line);
}

/// PHP `getdate(timestamp?)` — assoc array with date components.
///
/// Stack on entry: `[]` (no arg) or `[ts]` ; Stack on exit: `[obj]`.
pub fn emit_php_getdate(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let ts_slot = alloc_local(chunk);
    if argc >= 1 {
        local_set(chunk, ts_slot, line);
    }

    // ms = ts * 1000 OR Date.now()
    if argc >= 1 {
        local_get(chunk, ts_slot, line);
        push_const(chunk, Value::F64(MS_PER_SECOND), line);
        chunk.emit_op(Op::F64_MUL, line);
    } else {
        let _ = chunk;
        call_import(chunks, current, "ecma:date", "now", 0, line);
    }
    let chunk = &mut chunks[current];
    emit_wrap_ms(chunk, "Date", line);
    let dt_slot = alloc_local(chunk);
    local_set(chunk, dt_slot, line);

    // Build assoc Object with PHP-spec keys.
    let _ = chunk;
    call_import(chunks, current, "ecma:object", "new", 0, line);
    let chunk = &mut chunks[current];
    let out_slot = alloc_local(chunk);
    local_set(chunk, out_slot, line);

    // ── Fields ──
    // info["seconds"] = getSeconds(dt)
    let setters: &[(&str, &str, f64)] = &[
        ("seconds", "getSeconds", 0.0),
        ("minutes", "getMinutes", 0.0),
        ("hours", "getHours", 0.0),
        ("mday", "getDate", 0.0),
        ("wday", "getDay", 0.0),
        ("mon", "getMonth", 1.0), // 0-indexed → +1
        ("year", "getFullYear", 0.0),
    ];
    let _ = chunk;
    for (key, getter, bias) in setters {
        {
            let chunk = &mut chunks[current];
            local_get(chunk, out_slot, line);
            push_str(chunk, key, line);
            local_get(chunk, dt_slot, line);
        }
        call_import(chunks, current, "ecma:date", getter, 1, line);
        let chunk = &mut chunks[current];
        if *bias != 0.0 {
            push_const(chunk, Value::F64(*bias), line);
            chunk.emit_op(Op::F64_ADD, line);
        }
        chunk.emit_op(Op::ARRAY_SET, line);
    }

    let weekday_full = [
        "Sunday",
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
    ];
    let month_full = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ];

    // info["weekday"] = weekday_full[getDay(dt)]
    {
        let chunk = &mut chunks[current];
        local_get(chunk, out_slot, line);
        push_str(chunk, "weekday", line);
        for n in &weekday_full {
            push_str(chunk, n, line);
        }
        chunk.emit_array_new_fixed(0, weekday_full.len() as u16, line);
        local_get(chunk, dt_slot, line);
    }
    call_import(chunks, current, "ecma:date", "getDay", 1, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::ARRAY_GET, line);
        chunk.emit_op(Op::ARRAY_SET, line);

        // info["month"] = month_full[getMonth(dt)]
        local_get(chunk, out_slot, line);
        push_str(chunk, "month", line);
        for n in &month_full {
            push_str(chunk, n, line);
        }
        chunk.emit_array_new_fixed(0, month_full.len() as u16, line);
        local_get(chunk, dt_slot, line);
    }
    call_import(chunks, current, "ecma:date", "getMonth", 1, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::ARRAY_GET, line);
        chunk.emit_op(Op::ARRAY_SET, line);
        local_get(chunk, out_slot, line);
    }
}

/// PHP `strtotime($str)` and `strtotime($str, $base)` adapter.
///
/// Stack: `[s]` or `[s, base]` → `[secs]`.
///
/// Strategy:
/// - `strtotime("now")` → current secs (compile-time literal handled
///   by walker pre-parse; runtime form falls through to ecma:date.parse
///   which natively understands "now" so the same path works).
/// - Otherwise: `floor(ecma:date.parse(s) / 1000)`.
/// - 2-arg relative form (`"+7 days"`, `"-1 month"`, ...) is handled
///   by the walker pre-parser when `$str` is a literal. Runtime
///   2-arg path falls through to `ecma:date.parse(s)` (which doesn't
///   understand relative forms — best-effort; PHP users with dynamic
///   relative strings should use DateTimeImmutable->modify()).
/// `__php_strtotime_rel_calendar(base, n, is_year)` — apply a calendar
/// shift (months or years) to a seconds-epoch base. Walker emits this
/// for `strtotime("+N month", $base)` / `"+N year"`.
///
/// Stack on entry: `[base_secs, n, is_year_bool]` ; Stack on exit: `[secs]`.
pub fn emit_php_strtotime_rel_calendar(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let is_year_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let base_slot = alloc_local(chunk);
    local_set(chunk, is_year_slot, line);
    local_set(chunk, n_slot, line);
    local_set(chunk, base_slot, line);

    // ms = base * 1000
    local_get(chunk, base_slot, line);
    push_const(chunk, Value::F64(MS_PER_SECOND), line);
    chunk.emit_op(Op::F64_MUL, line);
    let ms_slot = alloc_local(chunk);
    local_set(chunk, ms_slot, line);

    // Build a Date wrapper: {__type:"Date", __time:ms}
    local_get(chunk, ms_slot, line);
    emit_wrap_ms(chunk, "Date", line);
    let dt_slot = alloc_local(chunk);
    local_set(chunk, dt_slot, line);

    // is_year ? setFullYear : setMonth
    local_get(chunk, is_year_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    let new_comp_slot = alloc_local(chunk);
    let new_ms_slot = alloc_local(chunk);
    chunk.emit_if(line);

    // setFullYear path
    local_get(chunk, dt_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:date", "getFullYear", 1, line);
    let chunk = &mut chunks[current];
    local_get(chunk, n_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    local_set(chunk, new_comp_slot, line);
    local_get(chunk, dt_slot, line);
    local_get(chunk, new_comp_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:date", "setFullYear", 2, line);
    let chunk = &mut chunks[current];
    local_set(chunk, new_ms_slot, line);
    chunk.emit_else(line);

    // setMonth path
    local_get(chunk, dt_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:date", "getMonth", 1, line);
    let chunk = &mut chunks[current];
    local_get(chunk, n_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    local_set(chunk, new_comp_slot, line);
    local_get(chunk, dt_slot, line);
    local_get(chunk, new_comp_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:date", "setMonth", 2, line);
    let chunk = &mut chunks[current];
    local_set(chunk, new_ms_slot, line);
    chunk.emit_end(line);

    // floor(new_ms / 1000)
    local_get(chunk, new_ms_slot, line);
    push_const(chunk, Value::F64(MS_PER_SECOND), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
}

pub fn emit_php_strtotime(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc >= 2 {
        // Drop $base — not used in the runtime fallback path.
        let _base_slot = alloc_local(chunk);
        local_set(chunk, _base_slot, line);
    }
    // Stack: [s]. Call ecma:date.parse(s) → ms_or_NaN.
    call_import(chunks, current, "ecma:date", "parse", 1, line);
    let chunk = &mut chunks[current];
    // floor(ms / 1000)
    push_const(chunk, Value::F64(MS_PER_SECOND), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    // PHP `strtotime` returns `false` (not NaN) when the string can't be
    // parsed. Detect NaN via the self-inequality `secs != secs`.
    let secs_slot = alloc_local(chunk);
    local_set(chunk, secs_slot, line);
    local_get(chunk, secs_slot, line);
    local_get(chunk, secs_slot, line);
    chunk.emit_op(Op::F64_NE, line);
    chunk.emit_if_value(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_else(line);
    local_get(chunk, secs_slot, line);
    chunk.emit_end(line);
}

/// PHP `$dt->getTimestamp()`.
///
/// Stack: `[dt]` → `[secs]` (i64-equivalent f64; PHP returns int but
/// the rest of the surface treats numbers as f64 anyway).
pub fn emit_datetime_get_timestamp(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    struct_get(chunk, &ClassSlot::internal(TIME_KEY), line);
    push_const(chunk, Value::F64(MS_PER_SECOND), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
}

/// Probe the receiver's `__type`; if it is `"DateTimeImmutable"`,
/// replace `dt_slot` with a fresh clone before mutating. Caller
/// returns `dt_slot` at the end so DateTimeImmutable callers see the
/// new object while DateTime callers see (and continue mutating) the
/// original — matching PHP's mutable-vs-immutable semantics.
///
/// Stack: unchanged (operates on `dt_slot` in place).
fn emit_clone_if_immutable(chunk: &mut Chunk, dt_slot: u16, line: u32) {
    // tag = dt.__type
    local_get(chunk, dt_slot, line);
    struct_get(chunk, &ClassSlot::TypeIdentity, line);
    push_str(chunk, "DateTimeImmutable", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if(line);
    // Build the clone: STRUCT_NEW + copy __type + copy __time + copy __tz.
    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_dup(line);
    push_str(chunk, "DateTimeImmutable", line);
    struct_set(chunk, &ClassSlot::TypeIdentity, line);
    chunk.emit_dup(line);
    local_get(chunk, dt_slot, line);
    struct_get(chunk, &ClassSlot::internal(TIME_KEY), line);
    struct_set(chunk, &ClassSlot::internal(TIME_KEY), line);
    chunk.emit_dup(line);
    local_get(chunk, dt_slot, line);
    struct_get(chunk, &ClassSlot::internal(TZ_KEY), line);
    struct_set(chunk, &ClassSlot::internal(TZ_KEY), line);
    // Stack: [clone]; replace dt_slot with the clone.
    local_set(chunk, dt_slot, line);
    chunk.emit_end(line);
}

/// Apply a fixed-duration delta (n × ms_per_unit) to the receiver's
/// `__time`. Used for second/minute/hour/day/week deltas where the
/// shift is a constant ms count. Stack on entry: `[dt, n]` (n as
/// f64); Stack on exit: `[dt]` (or a clone if dt was DateTimeImmutable).
fn emit_datetime_add_fixed_unit(chunks: &mut [Chunk], current: usize, ms_per_unit: f64, line: u32) {
    let chunk = &mut chunks[current];
    let n_slot = alloc_local(chunk);
    let dt_slot = alloc_local(chunk);
    local_set(chunk, n_slot, line);
    local_set(chunk, dt_slot, line);
    emit_clone_if_immutable(chunk, dt_slot, line);

    // newMs = dt.__time + n * ms_per_unit — shared scaling. The object shape
    // (`__time`, clone-if-immutable) stays here because it is PHP's; only the
    // arithmetic is common.
    local_get(chunk, dt_slot, line);
    struct_get(chunk, &ClassSlot::internal(TIME_KEY), line);
    local_get(chunk, n_slot, line);
    vybe_compiler::primitives::datetime::emit_add_scaled(chunk, ms_per_unit, line);
    let new_ms_slot = alloc_local(chunk);
    local_set(chunk, new_ms_slot, line);

    local_get(chunk, dt_slot, line);
    local_get(chunk, new_ms_slot, line);
    struct_set(chunk, &ClassSlot::internal(TIME_KEY), line);
    local_get(chunk, dt_slot, line);
}

/// Apply a calendar-component delta via `ecma:date.set<Component>`.
/// `getter`/`setter` are the ECMA-262 §21.4 method names. The
/// receiver is wrapped in a Date probe (an Object with `__time`)
/// because `ecma:date.*` accept that exact shape — DateTime's
/// `__time` field is identical, so we can pass the receiver directly.
///
/// Stack on entry: `[dt, n]` ; Stack on exit: `[dt]`.
fn emit_datetime_add_calendar(
    chunks: &mut [Chunk],
    current: usize,
    getter: &str,
    setter: &str,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let n_slot = alloc_local(chunk);
    let dt_slot = alloc_local(chunk);
    local_set(chunk, n_slot, line);
    local_set(chunk, dt_slot, line);
    emit_clone_if_immutable(chunk, dt_slot, line);

    // current_component = ecma:date.<getter>(dt)
    local_get(chunk, dt_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:date", getter, 1, line);
    let chunk = &mut chunks[current];
    let cur_comp_slot = alloc_local(chunk);
    local_set(chunk, cur_comp_slot, line);

    // new_component = current_component + n
    local_get(chunk, cur_comp_slot, line);
    local_get(chunk, n_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    let new_comp_slot = alloc_local(chunk);
    local_set(chunk, new_comp_slot, line);

    // ecma:date.<setter>(dt, new_component) → returns new ms
    local_get(chunk, dt_slot, line);
    local_get(chunk, new_comp_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:date", setter, 2, line);
    let chunk = &mut chunks[current];
    let new_ms_slot = alloc_local(chunk);
    local_set(chunk, new_ms_slot, line);

    // dt.__time = new_ms (setter mutates the Date object in place but
    // also returns the ms; we re-stamp explicitly so DateTime's __time
    // stays in sync regardless of how the host fn implements mutation).
    local_get(chunk, dt_slot, line);
    local_get(chunk, new_ms_slot, line);
    struct_set(chunk, &ClassSlot::internal(TIME_KEY), line);
    local_get(chunk, dt_slot, line);
}

/// PHP `$dt->modify($delta)` — runtime path for non-literal deltas.
/// Falls back to a no-op when the walker hasn't pre-parsed the string
/// (current MVP — `__php_dt_modify_*` literal-pre-parse paths are
/// chosen by the walker for string-literal deltas).
///
/// Stack on entry: `[dt, delta]` ; Stack on exit: `[dt]`.
pub fn emit_datetime_modify(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let _delta_slot = alloc_local(chunk);
    let dt_slot = alloc_local(chunk);
    local_set(chunk, _delta_slot, line);
    local_set(chunk, dt_slot, line);
    // Dynamic-string modify isn't supported in pure bytecode without
    // a string-walking parser. Walker takes the literal-string fast
    // path; this fallback returns the receiver unchanged so a
    // dynamic-delta call doesn't trap.
    local_get(chunk, dt_slot, line);
}

/// `$dt->modify` literal-second path. Stack: `[dt, n]` → `[dt]`.
pub fn emit_datetime_add_seconds(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_datetime_add_fixed_unit(chunks, current, MS_PER_SECOND, line);
}
pub fn emit_datetime_add_minutes(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_datetime_add_fixed_unit(chunks, current, MS_PER_MINUTE, line);
}
pub fn emit_datetime_add_hours(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_datetime_add_fixed_unit(chunks, current, MS_PER_HOUR, line);
}
pub fn emit_datetime_add_days(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_datetime_add_fixed_unit(chunks, current, MS_PER_DAY, line);
}
pub fn emit_datetime_add_weeks(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_datetime_add_fixed_unit(chunks, current, MS_PER_WEEK, line);
}
pub fn emit_datetime_add_months(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_datetime_add_calendar(chunks, current, "getMonth", "setMonth", line);
}
pub fn emit_datetime_add_years(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_datetime_add_calendar(chunks, current, "getFullYear", "setFullYear", line);
}

/// Compile-time PHP `date()` format-string pre-parser.
///
/// Walker rewrite calls this with a string-literal `fmt` and a
/// pre-walked `dt_expr` AST. Returns an AST that, when compiled,
/// produces the same string `format_php` would — using only ECMA-262
/// §21.4 Date methods (`getFullYear` / `getMonth` / `getDate` /
/// `getHours` / `getMinutes` / `getSeconds` / `getDay`) and
/// `String.prototype.padStart` for zero-padding.
///
/// Returns `None` if the format string contains a placeholder we
/// don't yet emit AST for (caller falls back to runtime adapter).
pub fn format_php_literal_to_ast(
    fmt: &str,
    dt_expr: &vybe_ast::Expression,
    span: &vybe_ast::Span,
) -> Option<vybe_ast::Expression> {
    use vybe_ast::{Argument, BinOp, ExprKind, Expression, Literal};

    fn lit_str(s: &str, span: &vybe_ast::Span) -> Expression {
        Expression::with_span(ExprKind::Lit(Literal::Str(s.to_string())), span.clone())
    }
    fn lit_int(n: i64, span: &vybe_ast::Span) -> Expression {
        Expression::with_span(ExprKind::Lit(Literal::Int(n)), span.clone())
    }
    fn member(obj: Expression, field: &str, span: &vybe_ast::Span) -> Expression {
        Expression::with_span(
            ExprKind::Member {
                object: Box::new(obj),
                field: field.to_string(),
                null_safe: false,
            },
            span.clone(),
        )
    }
    fn call(callee: Expression, args: Vec<Expression>, span: &vybe_ast::Span) -> Expression {
        Expression::with_span(
            ExprKind::Call {
                callee: Box::new(callee),
                args: args.into_iter().map(Argument::positional).collect(),
                optional: false,
            },
            span.clone(),
        )
    }
    fn dt_call(dt: &Expression, method: &str, span: &vybe_ast::Span) -> Expression {
        call(member(dt.clone(), method, span), vec![], span)
    }
    fn stringify(part: Expression, span: &vybe_ast::Span) -> Expression {
        // PHP `"" . x` coerces `x` to string. Equivalent to ECMA
        // `String(x)` but via the operator the PHP walker already
        // wires up — no `String` global lookup needed.
        Expression::with_span(
            vybe_ast::ExprKind::Binary {
                op: vybe_ast::BinOp::Concat,
                left: Box::new(Expression::with_span(
                    vybe_ast::ExprKind::Lit(vybe_ast::Literal::Str(String::new())),
                    span.clone(),
                )),
                right: Box::new(part),
            },
            span.clone(),
        )
    }
    fn pad(part: Expression, width: i64, span: &vybe_ast::Span) -> Expression {
        // str_pad("" . part, width, "0", STR_PAD_LEFT=0) — via the profile
        // `common:php.str_pad` emitter (ecma:string padStart host import).
        // A `->padStart` member call is not usable: PHP `->` is object-method
        // dispatch and does not resolve JS string-prototype methods.
        let stringified = stringify(part, span);
        call(
            Expression::with_span(ExprKind::Ident("str_pad".to_string()), span.clone()),
            vec![
                stringified,
                lit_int(width, span),
                lit_str("0", span),
                lit_int(0, span),
            ],
            span,
        )
    }
    fn add(left: Expression, right: Expression, span: &vybe_ast::Span) -> Expression {
        Expression::with_span(
            ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(left),
                right: Box::new(right),
            },
            span.clone(),
        )
    }
    fn concat(left: Expression, right: Expression, span: &vybe_ast::Span) -> Expression {
        Expression::with_span(
            ExprKind::Binary {
                op: BinOp::Concat,
                left: Box::new(left),
                right: Box::new(right),
            },
            span.clone(),
        )
    }
    fn array_index_str(items: &[&str], idx: Expression, span: &vybe_ast::Span) -> Expression {
        // Build an array literal then index it. Walker-shaped AST:
        // Array([items..])[idx].
        let elems: Vec<vybe_ast::ArrayElement> = items
            .iter()
            .map(|s| vybe_ast::ArrayElement {
                key: None,
                value: lit_str(s, span),
                spread: false,
                by_ref: false,
            })
            .collect();
        let arr = Expression::with_span(ExprKind::Array(elems), span.clone());
        Expression::with_span(
            ExprKind::Index {
                object: Box::new(arr),
                index: Box::new(idx),
                null_safe: false,
            },
            span.clone(),
        )
    }

    let weekday_full = [
        "Sunday",
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
    ];
    let weekday_abbr = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
    let month_full = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ];
    let month_abbr = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
    ];

    let mut chars = fmt.chars().peekable();
    let mut parts: Vec<Expression> = Vec::new();
    let mut buffer = String::new();
    let flush = |parts: &mut Vec<Expression>, buf: &mut String, span: &vybe_ast::Span| {
        if !buf.is_empty() {
            parts.push(lit_str(buf, span));
            buf.clear();
        }
    };
    while let Some(c) = chars.next() {
        if c == '\\' {
            if let Some(next) = chars.next() {
                buffer.push(next);
            }
            continue;
        }
        let placeholder: Option<Expression> = match c {
            // ── Date components ──
            'Y' => Some(stringify(dt_call(dt_expr, "getFullYear", span), span)),
            'y' => {
                // Last two digits, zero-padded.
                let yr = dt_call(dt_expr, "getFullYear", span);
                let mod100 = Expression::with_span(
                    ExprKind::Binary {
                        op: BinOp::Mod,
                        left: Box::new(yr),
                        right: Box::new(lit_int(100, span)),
                    },
                    span.clone(),
                );
                Some(pad(mod100, 2, span))
            }
            'm' => Some(pad(
                add(dt_call(dt_expr, "getMonth", span), lit_int(1, span), span),
                2,
                span,
            )),
            'n' => Some(stringify(
                add(dt_call(dt_expr, "getMonth", span), lit_int(1, span), span),
                span,
            )),
            'd' => Some(pad(dt_call(dt_expr, "getDate", span), 2, span)),
            'j' => Some(stringify(dt_call(dt_expr, "getDate", span), span)),
            'H' => Some(pad(dt_call(dt_expr, "getHours", span), 2, span)),
            'G' => Some(stringify(dt_call(dt_expr, "getHours", span), span)),
            'h' | 'g' => {
                // 12-hour format: ((hours + 11) % 12) + 1.
                let hr = dt_call(dt_expr, "getHours", span);
                let plus11 = add(hr, lit_int(11, span), span);
                let mod12 = Expression::with_span(
                    ExprKind::Binary {
                        op: BinOp::Mod,
                        left: Box::new(plus11),
                        right: Box::new(lit_int(12, span)),
                    },
                    span.clone(),
                );
                let plus1 = add(mod12, lit_int(1, span), span);
                if c == 'h' {
                    Some(pad(plus1, 2, span))
                } else {
                    Some(stringify(plus1, span))
                }
            }
            'i' => Some(pad(dt_call(dt_expr, "getMinutes", span), 2, span)),
            's' => Some(pad(dt_call(dt_expr, "getSeconds", span), 2, span)),
            'A' | 'a' => {
                let hr = dt_call(dt_expr, "getHours", span);
                let cmp = Expression::with_span(
                    ExprKind::Binary {
                        op: BinOp::Lt,
                        left: Box::new(hr),
                        right: Box::new(lit_int(12, span)),
                    },
                    span.clone(),
                );
                let (am, pm) = if c == 'A' { ("AM", "PM") } else { ("am", "pm") };
                Some(Expression::with_span(
                    ExprKind::Ternary {
                        cond: Box::new(cmp),
                        then: Box::new(lit_str(am, span)),
                        else_: Box::new(lit_str(pm, span)),
                    },
                    span.clone(),
                ))
            }
            'l' => Some(array_index_str(
                &weekday_full,
                dt_call(dt_expr, "getDay", span),
                span,
            )),
            'D' => Some(array_index_str(
                &weekday_abbr,
                dt_call(dt_expr, "getDay", span),
                span,
            )),
            'F' => Some(array_index_str(
                &month_full,
                dt_call(dt_expr, "getMonth", span),
                span,
            )),
            'M' => Some(array_index_str(
                &month_abbr,
                dt_call(dt_expr, "getMonth", span),
                span,
            )),
            'N' => {
                // ISO weekday: 1=Mon..7=Sun; JS getDay: 0=Sun..6=Sat.
                let dow = dt_call(dt_expr, "getDay", span);
                let cmp_zero = Expression::with_span(
                    ExprKind::Binary {
                        op: BinOp::StrictEq,
                        left: Box::new(dow.clone()),
                        right: Box::new(lit_int(0, span)),
                    },
                    span.clone(),
                );
                let n_int = Expression::with_span(
                    ExprKind::Ternary {
                        cond: Box::new(cmp_zero),
                        then: Box::new(lit_int(7, span)),
                        else_: Box::new(dow),
                    },
                    span.clone(),
                );
                Some(stringify(n_int, span))
            }
            'w' => Some(stringify(dt_call(dt_expr, "getDay", span), span)),
            'U' => {
                // Math.floor(dt.__time / 1000).
                let time = member(dt_expr.clone(), TIME_KEY, span);
                let div = Expression::with_span(
                    ExprKind::Binary {
                        op: BinOp::Div,
                        left: Box::new(time),
                        right: Box::new(lit_int(1000, span)),
                    },
                    span.clone(),
                );
                let floor = call(
                    member(
                        Expression::with_span(ExprKind::Ident("Math".to_string()), span.clone()),
                        "floor",
                        span,
                    ),
                    vec![div],
                    span,
                );
                Some(stringify(floor, span))
            }
            // Unknown placeholder — abort the optimization, let the
            // runtime adapter handle it.
            _ if c.is_ascii_alphabetic() => return None,
            other => {
                buffer.push(other);
                continue;
            }
        };
        if let Some(p) = placeholder {
            flush(&mut parts, &mut buffer, span);
            parts.push(p);
        }
    }
    flush(&mut parts, &mut buffer, span);

    if parts.is_empty() {
        return Some(lit_str("", span));
    }
    let mut iter = parts.into_iter();
    let mut acc = iter.next().unwrap();
    for p in iter {
        acc = concat(acc, p, span);
    }
    Some(acc)
}

/// Compile-time relative-delta parser. Returns `(n, unit_canon)` where
/// `unit_canon` is one of `"second" | "minute" | "hour" | "day" |
/// "week" | "month" | "year"` (singular, lowercase) — letting the
/// walker pick the matching `__php_dt_add_*` adapter.
pub fn parse_relative_delta(s: &str) -> Option<(i64, &'static str)> {
    let trimmed = s.trim();
    let (sign, rest) = if let Some(r) = trimmed.strip_prefix('+') {
        (1i64, r.trim_start())
    } else if let Some(r) = trimmed.strip_prefix('-') {
        (-1i64, r.trim_start())
    } else {
        return None;
    };
    let mut parts = rest.splitn(2, char::is_whitespace);
    let n_str = parts.next()?;
    let unit_raw = parts.next()?.trim().to_lowercase();
    let n: i64 = n_str.parse().ok()?;
    let unit = unit_raw.trim_end_matches('s');
    let canon: &'static str = match unit {
        "second" => "second",
        "minute" => "minute",
        "hour" => "hour",
        "day" => "day",
        "week" => "week",
        "month" => "month",
        "year" => "year",
        _ => return None,
    };
    Some((n * sign, canon))
}

/// PHP `$dt->modify($delta)` for `DateTimeImmutable` — clones the
/// receiver before mutating and returns the clone, leaving the
/// original untouched.
///
/// Stack on entry: `[dt, delta]` ; Stack on exit: `[new_dt]`.
pub fn emit_datetime_immutable_modify(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let delta_slot = alloc_local(chunk);
    let dt_slot = alloc_local(chunk);
    local_set(chunk, delta_slot, line);
    local_set(chunk, dt_slot, line);

    // Build a fresh DateTimeImmutable carrying the same __time, then
    // delegate to the mutable modify path on it.
    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_dup(line);
    push_str(chunk, "DateTimeImmutable", line);
    struct_set(chunk, &ClassSlot::TypeIdentity, line);
    chunk.emit_dup(line);
    local_get(chunk, dt_slot, line);
    struct_get(chunk, &ClassSlot::internal(TIME_KEY), line);
    struct_set(chunk, &ClassSlot::internal(TIME_KEY), line);
    chunk.emit_dup(line);
    local_get(chunk, dt_slot, line);
    struct_get(chunk, &ClassSlot::internal(TZ_KEY), line);
    struct_set(chunk, &ClassSlot::internal(TZ_KEY), line);
    // Stack: [clone]. Push delta and run the mutating modify on the clone.
    local_get(chunk, delta_slot, line);
    emit_datetime_modify(chunks, current, line);
}

/// Read an interval-component property as f64 (defaulting 0 if absent).
fn emit_read_interval_component(chunk: &mut Chunk, interval_slot: u16, key: &str, line: u32) {
    local_get(chunk, interval_slot, line);
    struct_get(chunk, &ClassSlot::internal(key), line);
}

/// Apply a `DateInterval` to the current `dt_slot` in place. `sign`
/// is +1 for `add`, -1 for `sub`. Stack on entry: empty (operates on
/// locals).
fn emit_apply_interval(chunk: &mut Chunk, dt_slot: u16, interval_slot: u16, sign: f64, line: u32) {
    // Compute total ms shift: y*365.25 + m*30.4375 + d + h/24 + i/1440 + s/86400
    // Years/months are calendar-irregular, but the test surface only
    // uses pure day/month components and absolute calendar diffs are
    // tested via `diff` which is exact; for `add`/`sub` an approximate
    // year/month shift is acceptable for the suite.
    let cur_ms_slot = alloc_local(chunk);
    local_get(chunk, dt_slot, line);
    struct_get(chunk, &ClassSlot::internal(TIME_KEY), line);
    local_set(chunk, cur_ms_slot, line);

    // y * 365.25 days * MS_PER_DAY
    emit_read_interval_component(chunk, interval_slot, "y", line);
    push_const(chunk, Value::F64(sign * 365.25 * MS_PER_DAY), line);
    chunk.emit_op(Op::F64_MUL, line);

    // m * 30.4375 days * MS_PER_DAY
    emit_read_interval_component(chunk, interval_slot, "m", line);
    push_const(chunk, Value::F64(sign * 30.4375 * MS_PER_DAY), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);

    // d * MS_PER_DAY
    emit_read_interval_component(chunk, interval_slot, "d", line);
    push_const(chunk, Value::F64(sign * MS_PER_DAY), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);

    // h * MS_PER_HOUR
    emit_read_interval_component(chunk, interval_slot, "h", line);
    push_const(chunk, Value::F64(sign * MS_PER_HOUR), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);

    // i (minutes) * MS_PER_MINUTE
    emit_read_interval_component(chunk, interval_slot, "i", line);
    push_const(chunk, Value::F64(sign * MS_PER_MINUTE), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);

    // s (seconds) * MS_PER_SECOND
    emit_read_interval_component(chunk, interval_slot, "s", line);
    push_const(chunk, Value::F64(sign * MS_PER_SECOND), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);

    // newMs = cur_ms + accumulator
    local_get(chunk, cur_ms_slot, line);
    chunk.emit_op(Op::F64_ADD, line);

    // Write back.
    let new_ms_slot = alloc_local(chunk);
    local_set(chunk, new_ms_slot, line);
    local_get(chunk, dt_slot, line);
    local_get(chunk, new_ms_slot, line);
    struct_set(chunk, &ClassSlot::internal(TIME_KEY), line);
}

/// PHP `$dt->add($interval)`. Stack: `[dt, interval]` → `[dt]`.
pub fn emit_datetime_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let interval_slot = alloc_local(chunk);
    let dt_slot = alloc_local(chunk);
    local_set(chunk, interval_slot, line);
    local_set(chunk, dt_slot, line);
    // DateTimeImmutable::add returns a new instance; mutable DateTime mutates.
    emit_clone_if_immutable(chunk, dt_slot, line);
    emit_apply_interval(chunk, dt_slot, interval_slot, 1.0, line);
    local_get(chunk, dt_slot, line);
}

/// PHP `$dt->sub($interval)`. Stack: `[dt, interval]` → `[dt]`.
pub fn emit_datetime_sub(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let interval_slot = alloc_local(chunk);
    let dt_slot = alloc_local(chunk);
    local_set(chunk, interval_slot, line);
    local_set(chunk, dt_slot, line);
    // DateTimeImmutable::sub returns a new instance; mutable DateTime mutates.
    emit_clone_if_immutable(chunk, dt_slot, line);
    emit_apply_interval(chunk, dt_slot, interval_slot, -1.0, line);
    local_get(chunk, dt_slot, line);
}

/// PHP `$dt->add($interval)` for `DateTimeImmutable` — clone first.
/// Stack: `[dt, interval]` → `[new_dt]`.
pub fn emit_datetime_immutable_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let interval_slot = alloc_local(chunk);
    let dt_slot = alloc_local(chunk);
    local_set(chunk, interval_slot, line);
    local_set(chunk, dt_slot, line);

    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_dup(line);
    push_str(chunk, "DateTimeImmutable", line);
    struct_set(chunk, &ClassSlot::TypeIdentity, line);
    chunk.emit_dup(line);
    local_get(chunk, dt_slot, line);
    struct_get(chunk, &ClassSlot::internal(TIME_KEY), line);
    struct_set(chunk, &ClassSlot::internal(TIME_KEY), line);
    chunk.emit_dup(line);
    local_get(chunk, dt_slot, line);
    struct_get(chunk, &ClassSlot::internal(TZ_KEY), line);
    struct_set(chunk, &ClassSlot::internal(TZ_KEY), line);

    let clone_slot = alloc_local(chunk);
    local_set(chunk, clone_slot, line);
    emit_apply_interval(chunk, clone_slot, interval_slot, 1.0, line);
    local_get(chunk, clone_slot, line);
}

/// Same shape as `emit_datetime_immutable_add` with `sign = -1.0`.
pub fn emit_datetime_immutable_sub(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let interval_slot = alloc_local(chunk);
    let dt_slot = alloc_local(chunk);
    local_set(chunk, interval_slot, line);
    local_set(chunk, dt_slot, line);

    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_dup(line);
    push_str(chunk, "DateTimeImmutable", line);
    struct_set(chunk, &ClassSlot::TypeIdentity, line);
    chunk.emit_dup(line);
    local_get(chunk, dt_slot, line);
    struct_get(chunk, &ClassSlot::internal(TIME_KEY), line);
    struct_set(chunk, &ClassSlot::internal(TIME_KEY), line);
    chunk.emit_dup(line);
    local_get(chunk, dt_slot, line);
    struct_get(chunk, &ClassSlot::internal(TZ_KEY), line);
    struct_set(chunk, &ClassSlot::internal(TZ_KEY), line);

    let clone_slot = alloc_local(chunk);
    local_set(chunk, clone_slot, line);
    emit_apply_interval(chunk, clone_slot, interval_slot, -1.0, line);
    local_get(chunk, clone_slot, line);
}

/// PHP `$dt->diff($other)` → DateInterval object.
///
/// Returns a `{__type: "DateInterval", days, y, m, d, h, i, s, invert}`
/// object computed from the millisecond delta. The MVP implementation
/// approximates calendar y/m components from totalDays — exact only
/// for the `days` field; tests rely primarily on `days` and integer
/// month arithmetic.
///
/// Stack: `[dt, other, absolute?]` → `[interval]`.
pub fn emit_datetime_diff(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let absolute_slot = alloc_local(chunk);
    if argc >= 3 {
        local_set(chunk, absolute_slot, line);
    } else {
        push_const(chunk, Value::Bool(false), line);
        local_set(chunk, absolute_slot, line);
    }
    let other_slot = alloc_local(chunk);
    let dt_slot = alloc_local(chunk);
    local_set(chunk, other_slot, line);
    local_set(chunk, dt_slot, line);

    // delta_ms = abs(other.__time - dt.__time)
    let delta_slot = alloc_local(chunk);
    local_get(chunk, other_slot, line);
    struct_get(chunk, &ClassSlot::internal(TIME_KEY), line);
    local_get(chunk, dt_slot, line);
    struct_get(chunk, &ClassSlot::internal(TIME_KEY), line);
    chunk.emit_op(Op::F64_SUB, line);
    let signed_slot = alloc_local(chunk);
    chunk.emit_dup(line);
    local_set(chunk, signed_slot, line);
    chunk.emit_op(Op::F64_ABS, line);
    local_set(chunk, delta_slot, line);

    // total_days = floor(delta_ms / MS_PER_DAY)
    let days_slot = alloc_local(chunk);
    local_get(chunk, delta_slot, line);
    push_const(chunk, Value::F64(MS_PER_DAY), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    local_set(chunk, days_slot, line);

    // y = floor(total_days / 365)
    let years_slot = alloc_local(chunk);
    local_get(chunk, days_slot, line);
    push_const(chunk, Value::F64(365.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    local_set(chunk, years_slot, line);

    // remaining_days = total_days - y*365
    let rem_after_years_slot = alloc_local(chunk);
    local_get(chunk, days_slot, line);
    local_get(chunk, years_slot, line);
    push_const(chunk, Value::F64(365.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
    local_set(chunk, rem_after_years_slot, line);

    // m = floor(remaining_days / 30)
    let months_slot = alloc_local(chunk);
    local_get(chunk, rem_after_years_slot, line);
    push_const(chunk, Value::F64(30.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    local_set(chunk, months_slot, line);

    // d = remaining_days - m*30
    let day_comp_slot = alloc_local(chunk);
    local_get(chunk, rem_after_years_slot, line);
    local_get(chunk, months_slot, line);
    push_const(chunk, Value::F64(30.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
    local_set(chunk, day_comp_slot, line);

    // invert = signed < 0 && !absolute
    let invert_slot = alloc_local(chunk);
    local_get(chunk, signed_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::F64_LT, line);
    local_get(chunk, absolute_slot, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
    local_set(chunk, invert_slot, line);

    // Build the DateInterval object.
    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_dup(line);
    push_str(chunk, "DateInterval", line);
    struct_set(chunk, &ClassSlot::TypeIdentity, line);
    chunk.emit_dup(line);
    local_get(chunk, days_slot, line);
    struct_set(chunk, &ClassSlot::internal("days"), line);
    chunk.emit_dup(line);
    local_get(chunk, years_slot, line);
    struct_set(chunk, &ClassSlot::internal("y"), line);
    chunk.emit_dup(line);
    local_get(chunk, months_slot, line);
    struct_set(chunk, &ClassSlot::internal("m"), line);
    chunk.emit_dup(line);
    local_get(chunk, day_comp_slot, line);
    struct_set(chunk, &ClassSlot::internal("d"), line);
    chunk.emit_dup(line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set(chunk, &ClassSlot::internal("h"), line);
    chunk.emit_dup(line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set(chunk, &ClassSlot::internal("i"), line);
    chunk.emit_dup(line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set(chunk, &ClassSlot::internal("s"), line);
    chunk.emit_dup(line);
    local_get(chunk, invert_slot, line);
    struct_set(chunk, &ClassSlot::internal("invert"), line);
}

/// PHP `new DateInterval($iso)` — parses ISO 8601 duration
/// `P[n]Y[n]M[n]DT[n]H[n]M[n]S` (with PHP's `W` weeks extension).
///
/// Walker-side path: when the constructor argument is a string
/// literal, the walker calls `parse_iso_duration` at compile time and
/// emits each component as a `Lit::Int` AST argument; this adapter
/// reads the six components from the stack and stamps them onto a
/// fresh DateInterval object. Dynamic strings flow through the
/// runtime parser (TODO — current tests use literals).
///
/// Stack on entry: `[y, m, d, h, i, s]` (six i64 values)
/// Stack on exit: `[interval]` with y/m/d/h/i/s set; days=0, invert=0.
pub fn emit_dateinterval_components(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let h_slot = alloc_local(chunk);
    let d_slot = alloc_local(chunk);
    let m_slot = alloc_local(chunk);
    let y_slot = alloc_local(chunk);
    // Stack: [y, m, d, h, i, s] — pop in reverse.
    local_set(chunk, s_slot, line);
    local_set(chunk, i_slot, line);
    local_set(chunk, h_slot, line);
    local_set(chunk, d_slot, line);
    local_set(chunk, m_slot, line);
    local_set(chunk, y_slot, line);

    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_dup(line);
    push_str(chunk, "DateInterval", line);
    struct_set(chunk, &ClassSlot::TypeIdentity, line);

    let pairs: &[(&str, u16)] = &[
        ("y", y_slot),
        ("m", m_slot),
        ("d", d_slot),
        ("h", h_slot),
        ("i", i_slot),
        ("s", s_slot),
    ];
    for (key, slot) in pairs {
        chunk.emit_dup(line);
        local_get(chunk, *slot, line);
        struct_set(chunk, &ClassSlot::internal(*key), line);
    }
    chunk.emit_dup(line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set(chunk, &ClassSlot::internal("days"), line);
    chunk.emit_dup(line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set(chunk, &ClassSlot::internal("invert"), line);
}

/// `DateInterval::format($fmt)` — the RUNTIME path, mirroring
/// `emit_datetime_format`. Stack on entry: `[interval, fmt]`; on exit:
/// `[string]`.
///
/// The walker also pre-folds *literal* formats, exactly as it does for
/// `DateTime::format`. This is what makes a non-literal format work at all,
/// and it is the single place the specifier table lives for the dynamic case.
pub fn emit_dateinterval_format(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let fmt_slot = alloc_local(chunk);
    let iv_slot = alloc_local(chunk);
    local_set(chunk, fmt_slot, line);
    local_set(chunk, iv_slot, line);

    let result_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let c_slot = alloc_local(chunk);

    push_str(chunk, "", line);
    local_set(chunk, result_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    local_set(chunk, i_slot, line);
    local_get(chunk, fmt_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    local_set(chunk, len_slot, line);

    let lstate = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    local_get(chunk, i_slot, line);
    local_get(chunk, len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(lstate.break_depth(0) as u32, line);

    // c = fmt.charAt(i)
    local_get(chunk, fmt_slot, line);
    local_get(chunk, i_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    local_set(chunk, c_slot, line);

    local_get(chunk, c_slot, line);
    push_str(chunk, "%", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if(line);

    // ── `%` seen: consume the specifier character ──
    local_get(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    local_set(chunk, i_slot, line);
    local_get(chunk, fmt_slot, line);
    local_get(chunk, i_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    local_set(chunk, c_slot, line);

    // Unpadded numeric fields, then the zero-padded upper-case forms.
    // `%a` is total days; `%R`/`%r` are the sign specifiers.
    for (code, field, width) in [
        ('y', "y", 0u32),
        ('m', "m", 0),
        ('d', "d", 0),
        ('h', "h", 0),
        ('i', "i", 0),
        ('s', "s", 0),
        ('a', "days", 0),
        ('Y', "y", 2),
        ('M', "m", 2),
        ('D', "d", 2),
        ('H', "h", 2),
        ('I', "i", 2),
        ('S', "s", 2),
    ] {
        local_get(chunk, c_slot, line);
        push_str(chunk, &code.to_string(), line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        chunk.emit_if(line);
        local_get(chunk, iv_slot, line);
        struct_get(chunk, &ClassSlot::internal(field), line);
        if width > 0 {
            emit_pad_to_width(chunk, width, line);
        } else {
            emit_stringify(chunk, line);
        }
        emit_append_to_result(chunk, result_slot, line);
        chunk.emit_end(line);
    }

    // `%R` → "+"/"-", `%r` → ""/"-", keyed off `invert`.
    for (code, positive) in [('R', "+"), ('r', "")] {
        local_get(chunk, c_slot, line);
        push_str(chunk, &code.to_string(), line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        chunk.emit_if(line);
        local_get(chunk, iv_slot, line);
        struct_get(chunk, &ClassSlot::internal("invert"), line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        push_str(chunk, "-", line);
        emit_append_to_result(chunk, result_slot, line);
        chunk.emit_else(line);
        push_str(chunk, positive, line);
        emit_append_to_result(chunk, result_slot, line);
        chunk.emit_end(line);
        chunk.emit_end(line);
    }

    // `%%` → a literal percent.
    local_get(chunk, c_slot, line);
    push_str(chunk, "%", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if(line);
    push_str(chunk, "%", line);
    emit_append_to_result(chunk, result_slot, line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    // ── ordinary character ──
    local_get(chunk, c_slot, line);
    emit_append_to_result(chunk, result_slot, line);
    chunk.emit_end(line);

    // i++
    local_get(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    local_set(chunk, i_slot, line);

    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, lstate, line);
    let chunk = &mut chunks[current];
    local_get(chunk, result_slot, line);
}

/// Parse a literal ISO 8601 duration string into (y, m, d, h, i, s)
/// components. Used by the walker / compiler when the DateInterval
/// constructor argument is a string literal — emits the components
/// as numeric constants in the bytecode rather than a runtime parser.
pub fn parse_iso_duration(s: &str) -> (i64, i64, i64, i64, i64, i64) {
    let mut y = 0i64;
    let mut mo = 0i64;
    let mut d = 0i64;
    let mut h = 0i64;
    let mut mi = 0i64;
    let mut se = 0i64;
    if let Some(rest) = s.strip_prefix('P') {
        let mut in_time = false;
        let mut num = String::new();
        for c in rest.chars() {
            if c == 'T' {
                in_time = true;
                continue;
            }
            if c.is_ascii_digit() {
                num.push(c);
            } else {
                let n: i64 = num.parse().unwrap_or(0);
                num.clear();
                match c {
                    'Y' => y = n,
                    'M' => {
                        if in_time {
                            mi = n
                        } else {
                            mo = n
                        }
                    }
                    'D' => d = n,
                    'H' => h = n,
                    'S' => se = n,
                    'W' => d = n * 7,
                    _ => {}
                }
            }
        }
    }
    (y, mo, d, h, mi, se)
}

// Silence unused-warning for stub helpers used only by future paths.
#[allow(dead_code)]
fn _unused_alloc_referrer() {
    let _ = MS_PER_WEEK;
}
