//! .NET `System.TimeZoneInfo` — bytecode-only, over the shared tzdb.
//!
//! The type did not exist here at all: `TimeZoneInfo` appeared nowhere in
//! `platforms/dotnet`, so `TimeZoneInfo.Utc.Id` answered `undefined`.
//!
//! Nothing about zones is reimplemented. Every answer comes from
//! `primitives/datetime.rs`, which is the one place that talks to
//! `ecma:intl/timezone` — the same tzdb PHP's `DateTimeZone` and Java's
//! `TimeZone` read. This file is the .NET SPELLING on top of it: what the
//! members are called, which of them are static, and .NET's rule for what
//! `DateTimeKind` means when converting.

use std::sync::Arc;
use vybe_compiler::primitives::datetime as shared_datetime;
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use super::{datetime_adapter, timespan_adapter};

const TYPE_KEY: &str = "__type";
/// The tzdb identifier the object stands for. Every member reads it back from
/// here rather than closing over it, so a zone that arrives as a VALUE (from
/// `FindSystemTimeZoneById`) behaves exactly like `Utc` or `Local`.
const ZONE_KEY: &str = "__zone";
/// Half a year in ms — the step to the OPPOSITE season, used to find an
/// instant at which the zone is on standard time.
const HALF_YEAR_MS: f64 = 182.0 * 86_400_000.0;

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        _ => panic!("push_const: no WASM-compliant encoding for {val:?}"),
    }
}

/// `[obj, value] → []`
fn set_field(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, idx, line);
}

/// `[] → [value]`
fn get_field_from(chunk: &mut Chunk, obj_slot: u16, key: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, idx, line);
}

/// Both spellings of one field. A case-insensitive front end folds the member
/// name, so a PascalCase-only field is invisible to it.
fn set_field_both_spellings(chunk: &mut Chunk, obj_slot: u16, value_slot: u16, key: &str, line: u32) {
    for spelling in [key.to_string(), key.to_ascii_lowercase()] {
        chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
        set_field(chunk, &spelling, line);
    }
}

/// The zone's offset in SECONDS EAST at the instant in `ms_slot`.
/// Stack: `[] → [seconds_east]`
fn emit_offset_seconds(chunk: &mut Chunk, zone_slot: u16, ms_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, zone_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    shared_datetime::emit_zone_offset_seconds(chunk, line);
}

/// The zone's offset in MILLISECONDS EAST at the instant in `ms_slot`.
fn emit_offset_ms(chunk: &mut Chunk, zone_slot: u16, ms_slot: u16, line: u32) {
    emit_offset_seconds(chunk, zone_slot, ms_slot, line);
    push_const(chunk, Value::F64(1000.0), line);
    chunk.emit_op(Op::F64_MUL, line);
}

/// Build the `TimeZoneInfo` object for the tzdb identifier on the stack.
/// Stack: `[zone_id] → [tz_obj]`
///
/// `BaseUtcOffset` is .NET's STANDARD-time offset, not the offset right now.
/// It is derived rather than guessed: ask tzdb whether the zone is on DST at
/// this instant, and if it is, read the offset half a year away — the opposite
/// season, which is standard time for any zone with a single DST period. No
/// "subtract an hour" assumption, which would be wrong for the 30-minute
/// zones.
fn emit_build_timezone(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let (zone_slot, now_slot, base_ms_slot, obj_slot, offset_slot) = {
        let chunk = &mut chunks[current];
        let zone_slot = chunk.alloc_scratch(5);
        chunk.emit_op_u16(Op::LOCAL_SET, zone_slot, line);
        (zone_slot, zone_slot + 1, zone_slot + 2, zone_slot + 3, zone_slot + 4)
    };

    let now_idx = chunks[current].add_import("ecma:date", "now");
    chunks[current].emit_call(now_idx, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, now_slot, line);

    // base_ms = isDst(zone, now) ? now + half a year : now
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, now_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, base_ms_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, zone_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, now_slot, line);
    shared_datetime::emit_zone_is_dst(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, now_slot, line);
    push_const(chunk, Value::F64(HALF_YEAR_MS), line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, base_ms_slot, line);
    chunk.emit_end(line);

    let object_new = chunk.add_import("ecma:object", "new");
    chunk.emit_call(object_new, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    push_const(chunk, Value::String(Arc::from("TimeZoneInfo")), line);
    set_field(chunk, TYPE_KEY, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, zone_slot, line);
    set_field(chunk, ZONE_KEY, line);

    // `Id`, `StandardName` and `DisplayName` are all the tzdb identifier —
    // .NET's Windows-registry display strings are not something tzdb carries,
    // and inventing one would be worse than answering with the id.
    for key in ["Id", "StandardName", "DisplayName"] {
        set_field_both_spellings(chunk, obj_slot, zone_slot, key, line);
    }

    emit_offset_ms(chunk, zone_slot, base_ms_slot, line);
    timespan_adapter::emit_build_timespan(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, offset_slot, line);
    set_field_both_spellings(chunk, obj_slot, offset_slot, "BaseUtcOffset", line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

/// `TimeZoneInfo.Utc`.
pub fn emit_timezone_utc(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    push_const(&mut chunks[current], Value::String(Arc::from("UTC")), line);
    emit_build_timezone(chunks, current, line);
}

/// `TimeZoneInfo.Local` — the host environment's zone, ECMA-262
/// `SystemTimeZoneIdentifier`. One clock for every language, so a zone set
/// from PHP is observed here.
pub fn emit_timezone_local(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    shared_datetime::emit_system_zone(&mut chunks[current], line);
    emit_build_timezone(chunks, current, line);
}

/// `TimeZoneInfo.FindSystemTimeZoneById(id)` — canonicalised through tzdb, so
/// a Link name (`US/Eastern`) resolves to its primary identifier.
///
/// ⛔ An unknown id CANONICALISES TO NULL and the object is built on the null,
/// which is the honest answer here: .NET throws `TimeZoneNotFoundException`,
/// and throwing from this emitter would need the exception class the
/// synthesized hierarchy does not carry. The null propagates to `Id`.
pub fn emit_timezone_find_by_id(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    shared_datetime::emit_zone_canonicalize(&mut chunks[current], line);
    emit_build_timezone(chunks, current, line);
}

/// The UTC INSTANT a DateTime stands for.
///
/// The object's `__time` is its wall-clock reading expressed as though UTC —
/// that is what `Date.UTC(y, m, d, …)` of the components produces — so for
/// `Kind = Utc` it already IS the instant, and for anything else the zone's
/// offset has to come back off. .NET reads `Unspecified` as local time for
/// `ConvertTime`, which is why only `Utc` is special-cased.
///
/// Stack: `[] → [utc_ms]`
fn emit_utc_instant_from_datetime(chunk: &mut Chunk, dt_slot: u16, line: u32) {
    let local_zone_slot = chunk.alloc_scratch(2);
    let wall_ms_slot = local_zone_slot + 1;
    get_field_from(chunk, dt_slot, "__time", line);
    chunk.emit_op_u16(Op::LOCAL_SET, wall_ms_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, wall_ms_slot, line);

    get_field_from(chunk, dt_slot, "Kind", line);
    chunk.emit_string_const("Utc", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, wall_ms_slot, line);
    chunk.emit_end(line);

    // Not UTC: subtract the LOCAL zone's offset at that reading.
    get_field_from(chunk, dt_slot, "Kind", line);
    chunk.emit_string_const("Utc", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_XOR, line);
    chunk.emit_if(line);
    shared_datetime::emit_system_zone(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, local_zone_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, wall_ms_slot, line);
    emit_offset_ms(chunk, local_zone_slot, wall_ms_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_end(line);
}

/// Wall-clock reading of `utc_ms` in `zone`, as a DateTime with `Kind`.
/// Stack: `[] → [datetime]`
fn emit_reading_in_zone(
    chunks: &mut Vec<Chunk>,
    current: usize,
    zone_slot: u16,
    utc_ms_slot: u16,
    kind: &str,
    line: u32,
) {
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, utc_ms_slot, line);
    emit_offset_ms(chunk, zone_slot, utc_ms_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    datetime_adapter::emit_datetime_from_millis(chunks, current, line);
    let chunk = &mut chunks[current];
    for spelling in ["Kind", "kind"] {
        core_wasm::dup(chunk, line);
        push_const(chunk, Value::String(Arc::from(kind)), line);
        set_field(chunk, spelling, line);
    }
}

/// `TimeZoneInfo.ConvertTime(dt, destinationZone)`.
pub fn emit_timezone_convert_time(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let (zone_slot, dt_slot, utc_ms_slot) = {
        let chunk = &mut chunks[current];
        let tz_slot = chunk.alloc_scratch(4);
        let dt_slot = tz_slot + 1;
        let utc_ms_slot = tz_slot + 2;
        let zone_slot = tz_slot + 3;
        chunk.emit_op_u16(Op::LOCAL_SET, tz_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, dt_slot, line);
        get_field_from(chunk, tz_slot, ZONE_KEY, line);
        chunk.emit_op_u16(Op::LOCAL_SET, zone_slot, line);
        emit_utc_instant_from_datetime(chunk, dt_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, utc_ms_slot, line);
        (zone_slot, dt_slot, utc_ms_slot)
    };
    let _ = dt_slot;
    emit_reading_in_zone(chunks, current, zone_slot, utc_ms_slot, "Unspecified", line);
}

/// `TimeZoneInfo.ConvertTimeToUtc(dt)`.
pub fn emit_timezone_convert_to_utc(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let (zone_slot, utc_ms_slot) = {
        let chunk = &mut chunks[current];
        let dt_slot = chunk.alloc_scratch(3);
        let utc_ms_slot = dt_slot + 1;
        let zone_slot = dt_slot + 2;
        chunk.emit_op_u16(Op::LOCAL_SET, dt_slot, line);
        emit_utc_instant_from_datetime(chunk, dt_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, utc_ms_slot, line);
        chunk.emit_string_const("UTC", line);
        chunk.emit_op_u16(Op::LOCAL_SET, zone_slot, line);
        (zone_slot, utc_ms_slot)
    };
    emit_reading_in_zone(chunks, current, zone_slot, utc_ms_slot, "Utc", line);
}

/// `TimeZoneInfo.ConvertTimeFromUtc(dt, destinationZone)` — `dt` is read as a
/// UTC instant whatever its Kind says, which is .NET's own rule for this one.
pub fn emit_timezone_convert_from_utc(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let (zone_slot, utc_ms_slot) = {
        let chunk = &mut chunks[current];
        let tz_slot = chunk.alloc_scratch(4);
        let dt_slot = tz_slot + 1;
        let utc_ms_slot = tz_slot + 2;
        let zone_slot = tz_slot + 3;
        chunk.emit_op_u16(Op::LOCAL_SET, tz_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, dt_slot, line);
        get_field_from(chunk, tz_slot, ZONE_KEY, line);
        chunk.emit_op_u16(Op::LOCAL_SET, zone_slot, line);
        get_field_from(chunk, dt_slot, "__time", line);
        chunk.emit_op_u16(Op::LOCAL_SET, utc_ms_slot, line);
        (zone_slot, utc_ms_slot)
    };
    emit_reading_in_zone(chunks, current, zone_slot, utc_ms_slot, "Unspecified", line);
}

/// `zone.GetUtcOffset(dt)` — a TimeSpan, at that instant (DST included).
pub fn emit_timezone_get_utc_offset(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    {
        let chunk = &mut chunks[current];
        let dt_slot = chunk.alloc_scratch(4);
        let tz_slot = dt_slot + 1;
        let zone_slot = dt_slot + 2;
        let utc_ms_slot = dt_slot + 3;
        chunk.emit_op_u16(Op::LOCAL_SET, dt_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, tz_slot, line);
        get_field_from(chunk, tz_slot, ZONE_KEY, line);
        chunk.emit_op_u16(Op::LOCAL_SET, zone_slot, line);
        emit_utc_instant_from_datetime(chunk, dt_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, utc_ms_slot, line);
        emit_offset_ms(chunk, zone_slot, utc_ms_slot, line);
    }
    timespan_adapter::emit_build_timespan(chunks, current, line);
}

/// `zone.IsDaylightSavingTime(dt)`.
pub fn emit_timezone_is_dst(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let dt_slot = chunk.alloc_scratch(4);
    let tz_slot = dt_slot + 1;
    let zone_slot = dt_slot + 2;
    let utc_ms_slot = dt_slot + 3;
    chunk.emit_op_u16(Op::LOCAL_SET, dt_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, tz_slot, line);
    get_field_from(chunk, tz_slot, ZONE_KEY, line);
    chunk.emit_op_u16(Op::LOCAL_SET, zone_slot, line);
    emit_utc_instant_from_datetime(chunk, dt_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, utc_ms_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, zone_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, utc_ms_slot, line);
    shared_datetime::emit_zone_is_dst(chunk, line);
}
