//! Python `datetime` adapter — bytecode-only.
//!
//! Mirrors `languages/php/src/emitter/datetime_adapter.rs`: each `emit_*`
//! composes pre-existing `ecma:date.*` host fns (`UTC`, `getUTCFullYear`,
//! …) into the Python-shaped surface. No new host fns are registered.
//!
//! Layout: `{__type, __time, year, month, day, hour, minute, second,
//! microsecond}`. Python reads components as bare attributes (`d.year`,
//! not PHP's `->format('Y')`), and the walker's `desugar_member_reads`
//! turns those into subscripts — so storing each component as a property
//! makes every attribute read a plain `STRUCT_GET` with zero frontend
//! work. Only methods (`isoformat`, `weekday`, `replace`) need dispatch.
//!
//! `__time` is ms-since-epoch (UTC); `__type` distinguishes
//! `date`/`time`/`datetime` for `isinstance` and repr.
//!
//! Every construction site funnels through `emit_materialize`, so the two
//! conversions that differ from ECMA live in exactly one place:
//!   * `ecma:date.getUTCMonth` is 0-based, Python's `.month` is 1-based.
//!   * `getUTCDay` is Sunday=0, Python's `weekday()` is Monday=0.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;
use vybe_emitter::instructions::core_wasm;

pub const TYPE_KEY: &str = "__type";
pub const TIME_KEY: &str = "__time";

/// `__type` tags. These double as the `isinstance` discriminator.
pub const TYPE_DATE: &str = "date";
pub const TYPE_TIME: &str = "time";
pub const TYPE_DATETIME: &str = "datetime";
pub const TYPE_TIMEDELTA: &str = "timedelta";
pub const TYPE_TIMEZONE: &str = "timezone";

/// A `timedelta`'s whole duration, in microseconds. `days`/`seconds`/
/// `microseconds` are the normalized views Python exposes.
const US_KEY: &str = "__us";

const US_PER_SECOND: f64 = 1_000_000.0;
const US_PER_DAY: f64 = 86_400_000_000.0;

/// The component properties materialized onto every point-in-time value,
/// paired with the `ecma:date` getter that derives each from `__time`.
/// `month` is fixed up (+1) after the getter — see the module note.
const COMPONENTS: &[(&str, &str)] = &[
    ("year", "getUTCFullYear"),
    ("month", "getUTCMonth"),
    ("day", "getUTCDate"),
    ("hour", "getUTCHours"),
    ("minute", "getUTCMinutes"),
    ("second", "getUTCSeconds"),
];

fn struct_set(chunk: &mut Chunk, key: &str, line: u32) {
    let k = chunk.add_constant(vybe_bytecode::Value::String(std::sync::Arc::from(key)));
    chunk.emit_op_u16(Op::STRUCT_SET, k, line);
    chunk.emit_op(Op::DROP, line);
}

fn struct_get(chunk: &mut Chunk, key: &str, line: u32) {
    let k = chunk.add_constant(vybe_bytecode::Value::String(std::sync::Arc::from(key)));
    chunk.emit_op_u16(Op::STRUCT_GET, k, line);
}

/// Where a materialized value's `__type` comes from. Arithmetic is
/// type-preserving (`date + timedelta` is a `date`, not a `datetime`), so
/// it carries the left operand's tag through from a local rather than
/// hardcoding one.
enum Tag<'a> {
    Const(&'a str),
    Local(u16),
}

/// Wrap a ms timestamp in a fully-materialized value.
/// Stack: `[ms]` → `[obj]`.
pub fn emit_materialize(chunk: &mut Chunk, type_tag: &str, line: u32) {
    emit_materialize_tag(chunk, Tag::Const(type_tag), line);
}

fn emit_materialize_tag(chunk: &mut Chunk, tag: Tag, line: u32) {
    let ms = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, ms, line);

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);

    chunk.emit_dup(line);
    match tag {
        Tag::Const(s) => chunk.emit_string_const(s, line),
        Tag::Local(slot) => chunk.emit_op_u16(Op::LOCAL_GET, slot, line),
    }
    struct_set(chunk, TYPE_KEY, line);

    chunk.emit_dup(line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    struct_set(chunk, TIME_KEY, line);

    for (prop, getter) in COMPONENTS {
        chunk.emit_dup(line);
        chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
        let idx = chunk.add_import("ecma:date", *getter);
        chunk.emit_call(idx, 1, line);
        if *prop == "month" {
            // ECMA months are 0-based; Python's are 1-based.
            core_wasm::f64_const(chunk, line, 1.0);
            chunk.emit_op(Op::F64_ADD, line);
        }
        struct_set(chunk, prop, line);
    }

    // microsecond: Python's sub-second field. ms-resolution source, so
    // it is always a whole number of milliseconds worth of micros.
    chunk.emit_dup(line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    let get_ms = chunk.add_import("ecma:date", "getUTCMilliseconds");
    chunk.emit_call(get_ms, 1, line);
    core_wasm::f64_const(chunk, line, 1000.0);
    chunk.emit_op(Op::F64_MUL, line);
    struct_set(chunk, "microsecond", line);
}

/// Push `ecma:date.UTC(y, m-1, d, h, mi, s)` from six locals. Stack: `[]` → `[ms]`.
fn emit_utc_from_locals(chunk: &mut Chunk, slots: &[u16; 6], line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slots[0], line); // year
    chunk.emit_op_u16(Op::LOCAL_GET, slots[1], line); // month (1-based)
    core_wasm::f64_const(chunk, line, 1.0);
    chunk.emit_op(Op::F64_SUB, line); // → 0-based for ECMA
    for slot in &slots[2..] {
        chunk.emit_op_u16(Op::LOCAL_GET, *slot, line);
    }
    let utc = chunk.add_import("ecma:date", "UTC");
    chunk.emit_call(utc, 6, line);
}

/// The epoch date backing a bare `time` — Python's `time` carries no date,
/// and nothing reads the date half of one.
const EPOCH_DEFAULTS: [f64; 6] = [1970.0, 1.0, 1.0, 0.0, 0.0, 0.0];

/// Shared component constructor for `date`/`time`/`datetime`.
///
/// `first` is the component the first argument fills (0 = year for
/// `date`/`datetime`, 3 = hour for `time`), so one routine covers all
/// three arities. Arguments past the six components (`microsecond`,
/// `tzinfo`) are accepted and dropped — they only need to compile.
fn emit_components_new(chunk: &mut Chunk, argc: u8, first: usize, type_tag: &str, line: u32) {
    let base = chunk.alloc_scratch(6);
    let slots = [base, base + 1, base + 2, base + 3, base + 4, base + 5];

    // Defaults first: any component the call omits keeps these.
    for (i, slot) in slots.iter().enumerate() {
        core_wasm::f64_const(chunk, line, EPOCH_DEFAULTS[i]);
        chunk.emit_op_u16(Op::LOCAL_SET, *slot, line);
    }

    // Arguments were pushed left-to-right, so unwind from the last.
    for i in (0..argc as usize).rev() {
        match slots.get(first + i) {
            Some(slot) => chunk.emit_op_u16(Op::LOCAL_SET, *slot, line),
            None => chunk.emit_op(Op::DROP, line),
        }
    }

    emit_utc_from_locals(chunk, &slots, line);
    emit_materialize(chunk, type_tag, line);
}

/// `datetime.date(y, m, d)`. Stack: `[y, m, d]` → `[date]`.
pub fn emit_date_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_components_new(&mut chunks[current], argc, 0, TYPE_DATE, line);
}

/// `datetime.time(h, m?, s?)`. Stack: `[h, …]` → `[time]`.
pub fn emit_time_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_components_new(&mut chunks[current], argc, 3, TYPE_TIME, line);
}

/// `datetime.datetime(y, m, d, h?, mi?, s?)`. Stack: `[y, m, d, …]` → `[datetime]`.
pub fn emit_datetime_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_components_new(&mut chunks[current], argc, 0, TYPE_DATETIME, line);
}

/// Wrap a microsecond duration in a `timedelta`. Stack: `[us]` → `[obj]`.
///
/// Python normalizes a duration so `seconds` is 0–86399 and `days` carries
/// the rest, flooring toward negative infinity (`-timedelta(days=1)` is
/// `days=-1, seconds=0`) — which is exactly `f64.floor`'s rounding.
pub fn emit_wrap_timedelta(chunk: &mut Chunk, line: u32) {
    let us = chunk.alloc_scratch(1);
    let days = chunk.alloc_scratch(1);
    let rem = chunk.alloc_scratch(1);
    let secs = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, us, line);

    // days = floor(us / US_PER_DAY)
    chunk.emit_op_u16(Op::LOCAL_GET, us, line);
    core_wasm::f64_const(chunk, line, US_PER_DAY);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    chunk.emit_op_u16(Op::LOCAL_SET, days, line);

    // rem = us - days * US_PER_DAY
    chunk.emit_op_u16(Op::LOCAL_GET, us, line);
    chunk.emit_op_u16(Op::LOCAL_GET, days, line);
    core_wasm::f64_const(chunk, line, US_PER_DAY);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_SET, rem, line);

    // seconds = floor(rem / US_PER_SECOND)
    chunk.emit_op_u16(Op::LOCAL_GET, rem, line);
    core_wasm::f64_const(chunk, line, US_PER_SECOND);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    chunk.emit_op_u16(Op::LOCAL_SET, secs, line);

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_dup(line);
    chunk.emit_string_const(TYPE_TIMEDELTA, line);
    struct_set(chunk, TYPE_KEY, line);
    chunk.emit_dup(line);
    chunk.emit_op_u16(Op::LOCAL_GET, us, line);
    struct_set(chunk, US_KEY, line);
    chunk.emit_dup(line);
    chunk.emit_op_u16(Op::LOCAL_GET, days, line);
    struct_set(chunk, "days", line);
    chunk.emit_dup(line);
    chunk.emit_op_u16(Op::LOCAL_GET, secs, line);
    struct_set(chunk, "seconds", line);
    // microseconds = rem - seconds * US_PER_SECOND
    chunk.emit_dup(line);
    chunk.emit_op_u16(Op::LOCAL_GET, rem, line);
    chunk.emit_op_u16(Op::LOCAL_GET, secs, line);
    core_wasm::f64_const(chunk, line, US_PER_SECOND);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
    struct_set(chunk, "microseconds", line);
}

/// How many microseconds each `timedelta` parameter contributes, in the
/// order of CPython's signature (the walker hands us that exact order).
const TIMEDELTA_SCALE: [f64; 7] = [
    US_PER_DAY,          // days
    US_PER_SECOND,       // seconds
    1.0,                 // microseconds
    1_000.0,             // milliseconds
    60.0 * US_PER_SECOND,   // minutes
    3600.0 * US_PER_SECOND, // hours
    7.0 * US_PER_DAY,    // weeks
];

/// `datetime.timedelta(days=…, seconds=…, …)`. Stack: `[args…]` → `[obj]`.
pub fn emit_timedelta_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let base = chunk.alloc_scratch(7);

    for i in (0..argc as usize).rev() {
        if i < 7 {
            chunk.emit_op_u16(Op::LOCAL_SET, base + i as u16, line);
        } else {
            chunk.emit_op(Op::DROP, line);
        }
    }

    // Sum only the parameters actually supplied; the rest are zero.
    core_wasm::f64_const(chunk, line, 0.0);
    for i in 0..(argc as usize).min(7) {
        chunk.emit_op_u16(Op::LOCAL_GET, base + i as u16, line);
        core_wasm::f64_const(chunk, line, TIMEDELTA_SCALE[i]);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_ADD, line);
    }

    emit_wrap_timedelta(chunk, line);
}

/// `timedelta.total_seconds()` → float. Stack: `[obj]` → `[num]`.
pub fn emit_total_seconds(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    struct_get(chunk, US_KEY, line);
    core_wasm::f64_const(chunk, line, US_PER_SECOND);
    chunk.emit_op(Op::F64_DIV, line);
}

/// Wrap a `timedelta` offset in a `timezone`. Stack: `[timedelta]` → `[tz]`.
fn emit_wrap_timezone(chunk: &mut Chunk, line: u32) {
    let off = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, off, line);

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_dup(line);
    chunk.emit_string_const(TYPE_TIMEZONE, line);
    struct_set(chunk, TYPE_KEY, line);
    chunk.emit_dup(line);
    chunk.emit_op_u16(Op::LOCAL_GET, off, line);
    struct_set(chunk, "__offset", line);
}

/// `datetime.timezone(offset)` — a fixed UTC offset carrying the
/// `timedelta` it was built from. Stack: `[timedelta]` → `[obj]`.
pub fn emit_timezone_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    // `timezone(offset, name)` — the name is display-only.
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    emit_wrap_timezone(chunk, line);
}

/// `datetime.timezone.utc` — the zero offset.
pub fn emit_timezone_utc(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    core_wasm::f64_const(chunk, line, 0.0);
    emit_wrap_timedelta(chunk, line);
    emit_wrap_timezone(chunk, line);
}

/// `datetime.timedelta.resolution` — one microsecond.
pub fn emit_timedelta_resolution(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    core_wasm::f64_const(chunk, line, 1.0);
    emit_wrap_timedelta(chunk, line);
}

/// Days from 1970-01-01 to a proleptic-Gregorian date (Hinnant's
/// `days_from_civil`).
///
/// ECMA's Date constructor remaps years 0–99 onto 1900–1999 (§21.4.2.7),
/// so `date.min` (year 1) would silently come back as 1901 if built
/// through `ecma:date.UTC`. Computing the epoch offset here sidesteps the
/// constructor; the `ecma:date` getters carry no such remapping, so
/// `emit_materialize` still reads the components back correctly.
fn days_from_civil(y: i64, m: i64, d: i64) -> i64 {
    let y = if m <= 2 { y - 1 } else { y };
    let era = if y >= 0 { y } else { y - 399 } / 400;
    let yoe = y - era * 400;
    let mp = (m + 9) % 12;
    let doy = (153 * mp + 2) / 5 + d - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146_097 + doe - 719_468
}

/// Python ordinal 1 is 0001-01-01; the epoch is ordinal 719163.
const EPOCH_ORDINAL: f64 = 719_163.0;
const MS_PER_DAY: f64 = 86_400_000.0;

fn emit_literal_date(chunk: &mut Chunk, y: i64, m: i64, d: i64, tag: &str, line: u32) {
    core_wasm::f64_const(chunk, line, (days_from_civil(y, m, d) * 86_400_000) as f64);
    emit_materialize(chunk, tag, line);
}

/// `datetime.date.min` / `datetime.datetime.min` → 0001-01-01.
pub fn emit_date_min(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_literal_date(&mut chunks[current], 1, 1, 1, TYPE_DATE, line);
}

/// `datetime.date.max` / `datetime.datetime.max` → 9999-12-31.
pub fn emit_date_max(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_literal_date(&mut chunks[current], 9999, 12, 31, TYPE_DATE, line);
}

/// `date.toordinal()` → days since 0001-01-01, 1-based.
pub fn emit_toordinal(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    struct_get(chunk, TIME_KEY, line);
    core_wasm::f64_const(chunk, line, MS_PER_DAY);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    core_wasm::f64_const(chunk, line, EPOCH_ORDINAL);
    chunk.emit_op(Op::F64_ADD, line);
}

/// `date.fromordinal(n)`. Stack: `[n]` → `[date]`.
pub fn emit_fromordinal(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    core_wasm::f64_const(chunk, line, EPOCH_ORDINAL);
    chunk.emit_op(Op::F64_SUB, line);
    core_wasm::f64_const(chunk, line, MS_PER_DAY);
    chunk.emit_op(Op::F64_MUL, line);
    emit_materialize(chunk, TYPE_DATE, line);
}

/// `tz.utcoffset(dt)` → the `timedelta` the zone was built from. The
/// argument only matters to a real zone with DST, which a fixed offset
/// is not. `argc` counts the receiver, so it alone is the non-argument.
/// Stack: `[tz, arg…]` → `[timedelta]`.
pub fn emit_utcoffset(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    struct_get(chunk, "__offset", line);
}

/// Drop every argument past the first `keep`. Builtin `argc` counts only
/// real arguments; a value method's also counts its receiver.
fn drop_extra_args(chunk: &mut Chunk, argc: u8, keep: u8, line: u32) {
    for _ in keep..argc {
        chunk.emit_op(Op::DROP, line);
    }
}

/// `datetime.fromtimestamp(seconds)`. Stack: `[sec, tz?]` → `[datetime]`.
pub fn emit_fromtimestamp(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_extra_args(chunk, argc, 1, line);
    core_wasm::f64_const(chunk, line, MS_PER_SECOND);
    chunk.emit_op(Op::F64_MUL, line);
    emit_materialize(chunk, TYPE_DATETIME, line);
}

const MS_PER_SECOND: f64 = 1_000.0;

/// `datetime.timestamp()` → seconds as a float. Stack: `[dt]` → `[num]`.
pub fn emit_timestamp(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    struct_get(chunk, TIME_KEY, line);
    core_wasm::f64_const(chunk, line, MS_PER_SECOND);
    chunk.emit_op(Op::F64_DIV, line);
}

/// `date.fromisoformat(s)`. Stack: `[s]` → `[date]`.
pub fn emit_date_fromisoformat(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_extra_args(chunk, argc, 1, line);
    let parse = chunk.add_import("ecma:date", "parse");
    chunk.emit_call(parse, 1, line);
    emit_materialize(chunk, TYPE_DATE, line);
}

/// `datetime.fromisoformat(s)`, and `strptime(s, fmt)` — for the ISO-shaped
/// formats these carry, parsing is the same operation. Stack: `[s, …]` → `[datetime]`.
pub fn emit_datetime_fromisoformat(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_extra_args(chunk, argc, 1, line);
    let parse = chunk.add_import("ecma:date", "parse");
    chunk.emit_call(parse, 1, line);
    emit_materialize(chunk, TYPE_DATETIME, line);
}

/// `time.fromisoformat("12:30:45")` — a bare time is not a parseable
/// instant, so it is dated to the epoch first. Stack: `[s]` → `[time]`.
pub fn emit_time_fromisoformat(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_extra_args(chunk, argc, 1, line);
    let s = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, s, line);
    chunk.emit_string_const("1970-01-01T", line);
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    let concat = chunk.add_import("wasm:js-string", "concat");
    chunk.emit_call(concat, 2, line);
    let parse = chunk.add_import("ecma:date", "parse");
    chunk.emit_call(parse, 1, line);
    emit_materialize(chunk, TYPE_TIME, line);
}

/// `datetime.now()` / `datetime.utcnow()`. Stack: `[tz?]` → `[datetime]`.
pub fn emit_now(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_extra_args(chunk, argc, 0, line);
    let now = chunk.add_import("ecma:date", "now");
    chunk.emit_call(now, 0, line);
    emit_materialize(chunk, TYPE_DATETIME, line);
}

/// Truncate a ms timestamp to midnight. Stack: `[ms]` → `[ms]`.
fn emit_floor_to_day(chunk: &mut Chunk, line: u32) {
    core_wasm::f64_const(chunk, line, MS_PER_DAY);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    core_wasm::f64_const(chunk, line, MS_PER_DAY);
    chunk.emit_op(Op::F64_MUL, line);
}

/// `date.today()`. Stack: `[]` → `[date]`.
pub fn emit_today(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_extra_args(chunk, argc, 0, line);
    let now = chunk.add_import("ecma:date", "now");
    chunk.emit_call(now, 0, line);
    emit_floor_to_day(chunk, line);
    emit_materialize(chunk, TYPE_DATE, line);
}

/// Milliseconds since midnight for the value in `obj`. Stack: `[]` → `[ms]`.
fn emit_time_of_day_ms(chunk: &mut Chunk, obj: u16, line: u32) {
    let part = |chunk: &mut Chunk, prop: &str, scale: f64| {
        chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
        struct_get(chunk, prop, line);
        core_wasm::f64_const(chunk, line, scale);
        chunk.emit_op(Op::F64_MUL, line);
    };
    part(chunk, "hour", 3_600_000.0);
    part(chunk, "minute", 60_000.0);
    chunk.emit_op(Op::F64_ADD, line);
    part(chunk, "second", MS_PER_SECOND);
    chunk.emit_op(Op::F64_ADD, line);
}

/// `datetime.combine(date, time)`. Stack: `[d, t, tz?]` → `[datetime]`.
pub fn emit_combine(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_extra_args(chunk, argc, 2, line);
    let t = chunk.alloc_scratch(1);
    let d = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, t, line);
    chunk.emit_op_u16(Op::LOCAL_SET, d, line);

    chunk.emit_op_u16(Op::LOCAL_GET, d, line);
    struct_get(chunk, TIME_KEY, line);
    emit_floor_to_day(chunk, line);
    emit_time_of_day_ms(chunk, t, line);
    chunk.emit_op(Op::F64_ADD, line);
    emit_materialize(chunk, TYPE_DATETIME, line);
}

/// `datetime.date()` — the date half. Stack: `[dt]` → `[date]`.
pub fn emit_date_method(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    struct_get(chunk, TIME_KEY, line);
    emit_floor_to_day(chunk, line);
    emit_materialize(chunk, TYPE_DATE, line);
}

/// `datetime.time()` — the time half, dated to the epoch. Stack: `[dt]` → `[time]`.
pub fn emit_time_method(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let obj = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, obj, line);
    emit_time_of_day_ms(chunk, obj, line);
    emit_materialize(chunk, TYPE_TIME, line);
}

/// `date.timetuple()` → a `struct_time`-shaped value. Its `tm_*` fields are
/// plain properties, which is exactly what an attribute read wants.
/// Stack: `[value]` → `[obj]`.
pub fn emit_timetuple(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let obj = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, obj, line);

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    for (tm, prop) in [
        ("tm_year", "year"),
        ("tm_mon", "month"),
        ("tm_mday", "day"),
        ("tm_hour", "hour"),
        ("tm_min", "minute"),
        ("tm_sec", "second"),
    ] {
        chunk.emit_dup(line);
        chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
        struct_get(chunk, prop, line);
        struct_set(chunk, tm, line);
    }
    // tm_wday is Monday=0, like `weekday()`.
    chunk.emit_dup(line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    struct_get(chunk, TIME_KEY, line);
    let get_day = chunk.add_import("ecma:date", "getUTCDay");
    chunk.emit_call(get_day, 1, line);
    core_wasm::i32_const(chunk, line, 6);
    chunk.emit_op(Op::I32_ADD, line);
    core_wasm::i32_const(chunk, line, 7);
    chunk.emit_op(Op::I32_REM_S, line);
    struct_set(chunk, "tm_wday", line);
    chunk.emit_dup(line);
    core_wasm::i32_const(chunk, line, -1);
    struct_set(chunk, "tm_isdst", line);
}

/// Weekday (Monday=0) of a y/m/d already on the stack, via a ms round-trip.
/// Stack: `[y, m, d]` → `[num]`.
fn emit_weekday_of_ymd(chunk: &mut Chunk, line: u32) {
    let d = chunk.alloc_scratch(1);
    let m = chunk.alloc_scratch(1);
    let y = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, d, line);
    chunk.emit_op_u16(Op::LOCAL_SET, m, line);
    chunk.emit_op_u16(Op::LOCAL_SET, y, line);

    chunk.emit_op_u16(Op::LOCAL_GET, y, line);
    chunk.emit_op_u16(Op::LOCAL_GET, m, line);
    core_wasm::f64_const(chunk, line, 1.0);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_GET, d, line);
    let utc = chunk.add_import("ecma:date", "UTC");
    chunk.emit_call(utc, 3, line);
    let get_day = chunk.add_import("ecma:date", "getUTCDay");
    chunk.emit_call(get_day, 1, line);
    core_wasm::i32_const(chunk, line, 6);
    chunk.emit_op(Op::I32_ADD, line);
    core_wasm::i32_const(chunk, line, 7);
    chunk.emit_op(Op::I32_REM_S, line);
}

/// `calendar.weekday(y, m, d)` — Monday=0. Stack: `[y, m, d]` → `[num]`.
pub fn emit_cal_weekday(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_weekday_of_ymd(&mut chunks[current], line);
}

/// `calendar.isleap(y)`. Stack: `[y]` → `[bool]`.
pub fn emit_cal_isleap(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let y = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, y, line);

    let divides = |chunk: &mut Chunk, n: i32| {
        chunk.emit_op_u16(Op::LOCAL_GET, y, line);
        core_wasm::i32_const(chunk, line, n);
        chunk.emit_op(Op::I32_REM_S, line);
        chunk.emit_op(Op::I32_EQZ, line);
    };

    // (y % 4 == 0 && y % 100 != 0) || y % 400 == 0
    divides(chunk, 4);
    divides(chunk, 100);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
    divides(chunk, 400);
    chunk.emit_op(Op::I32_OR, line);
    // Python's `isleap` returns a real `bool`, so lift the i32 the same way
    // the comparison operators do under `materialize_bool_results`.
    vybe_emitter::ops::emit_i32_to_bool(chunk, line);
}

/// Days in month `m` of year `y`. Day 0 of the next month is the last day
/// of this one — the standard ECMA overflow trick. Stack: `[]` → `[num]`.
fn emit_days_in_month(chunk: &mut Chunk, y: u16, m: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, y, line);
    chunk.emit_op_u16(Op::LOCAL_GET, m, line);
    core_wasm::f64_const(chunk, line, 0.0);
    let utc = chunk.add_import("ecma:date", "UTC");
    chunk.emit_call(utc, 3, line);
    let get_date = chunk.add_import("ecma:date", "getUTCDate");
    chunk.emit_call(get_date, 1, line);
}

/// `calendar.monthrange(y, m)` → `(weekday_of_first, days_in_month)`.
/// Stack: `[y, m]` → `[tuple]`.
pub fn emit_cal_monthrange(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let m = chunks[current].alloc_scratch(1);
    let y = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, m, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, y, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, y, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, m, line);
    core_wasm::f64_const(&mut chunks[current], line, 1.0);
    emit_weekday_of_ymd(&mut chunks[current], line);
    emit_days_in_month(&mut chunks[current], y, m, line);

    // Python's monthrange returns a real tuple — built the one canonical way.
    vybe_emitter::tuples::emit_tuple(chunks, current, 2, line);
}

/// The arithmetic this adapter implements for its own values.
pub enum DtOp {
    Add,
    Sub,
    Mul,
}

/// `slot` holds an object carrying `key`. The `typeof` guard matters:
/// `STRUCT_GET` traps on a primitive, and these run on every `+`.
/// Stack: `[]` → `[i32]`.
fn emit_has_key(chunk: &mut Chunk, slot: u16, key: &str, line: u32) {
    let typeof_fn = chunk.add_import("ecma:value", "typeof");
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_call(typeof_fn, 1, line);
    chunk.emit_string_const("object", line);
    vybe_emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    struct_get(chunk, key, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_else(line);
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_end(line);
}

/// True when `slot` holds any value this adapter produced — a duration
/// (`__us`) or a point in time (`__time`). Presence of the property IS the
/// test; no name matching. Stack: `[]` → `[i32]`.
pub fn emit_is_datetime(chunk: &mut Chunk, slot: u16, line: u32) {
    emit_has_key(chunk, slot, US_KEY, line);
    emit_has_key(chunk, slot, TIME_KEY, line);
    chunk.emit_op(Op::I32_OR, line);
}

fn dt_op_code(op: &DtOp) -> Op {
    match op {
        DtOp::Add => Op::F64_ADD,
        DtOp::Sub => Op::F64_SUB,
        DtOp::Mul => Op::F64_MUL,
    }
}

/// `+`/`-`/`*` over this adapter's values, for the combinations Python
/// defines:
///
/// * `timedelta ∘ timedelta` → timedelta
/// * `timedelta * number`    → timedelta
/// * `point ± timedelta`     → the *left operand's* type (`date + timedelta`
///   is a `date`, never a `datetime`)
/// * `point − point`         → timedelta
///
/// Reads `a`/`b`; pushes the result. Only called once `emit_is_datetime(a)`
/// holds, so the left operand always carries one of the two keys.
pub fn emit_dt_binop(chunk: &mut Chunk, a: u16, b: u16, op: DtOp, line: u32) {
    emit_has_key(chunk, a, US_KEY, line);
    chunk.emit_if_value(line);
    {
        // a is a duration.
        chunk.emit_op_u16(Op::LOCAL_GET, a, line);
        struct_get(chunk, US_KEY, line);
        emit_has_key(chunk, b, US_KEY, line);
        chunk.emit_if_value(line);
        chunk.emit_op_u16(Op::LOCAL_GET, b, line);
        struct_get(chunk, US_KEY, line);
        chunk.emit_else(line);
        // `timedelta * 2` — a bare number scales the duration.
        chunk.emit_op_u16(Op::LOCAL_GET, b, line);
        chunk.emit_end(line);
        chunk.emit_op(dt_op_code(&op), line);
        emit_wrap_timedelta(chunk, line);
    }
    chunk.emit_else(line);
    {
        // a is a point in time.
        emit_has_key(chunk, b, US_KEY, line);
        chunk.emit_if_value(line);
        {
            // point ± timedelta → same type as `a`.
            let tag = chunk.alloc_scratch(1);
            chunk.emit_op_u16(Op::LOCAL_GET, a, line);
            struct_get(chunk, TYPE_KEY, line);
            chunk.emit_op_u16(Op::LOCAL_SET, tag, line);

            chunk.emit_op_u16(Op::LOCAL_GET, a, line);
            struct_get(chunk, TIME_KEY, line);
            chunk.emit_op_u16(Op::LOCAL_GET, b, line);
            struct_get(chunk, US_KEY, line);
            core_wasm::f64_const(chunk, line, MS_PER_SECOND);
            chunk.emit_op(Op::F64_DIV, line);
            chunk.emit_op(dt_op_code(&op), line);
            emit_materialize_tag(chunk, Tag::Local(tag), line);
        }
        chunk.emit_else(line);
        {
            // point − point → duration.
            chunk.emit_op_u16(Op::LOCAL_GET, a, line);
            struct_get(chunk, TIME_KEY, line);
            chunk.emit_op_u16(Op::LOCAL_GET, b, line);
            struct_get(chunk, TIME_KEY, line);
            chunk.emit_op(dt_op_code(&op), line);
            core_wasm::f64_const(chunk, line, MS_PER_SECOND);
            chunk.emit_op(Op::F64_MUL, line);
            emit_wrap_timedelta(chunk, line);
        }
        chunk.emit_end(line);
    }
    chunk.emit_end(line);
}

/// Unary `-` on a duration. Stack: `[td]` → `[td]`.
pub fn emit_dt_neg(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    struct_get(chunk, US_KEY, line);
    core_wasm::f64_const(chunk, line, -1.0);
    chunk.emit_op(Op::F64_MUL, line);
    emit_wrap_timedelta(chunk, line);
}

/// Relational compare of two of this adapter's values, by the instant or
/// duration they denote. Stack: `[]` → `[i32]`.
pub fn emit_dt_cmp(chunk: &mut Chunk, a: u16, b: u16, cmp: fn(&mut Chunk, u32), line: u32) {
    let key = |chunk: &mut Chunk, slot: u16| {
        emit_has_key(chunk, slot, US_KEY, line);
        chunk.emit_if_value(line);
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
        struct_get(chunk, US_KEY, line);
        chunk.emit_else(line);
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
        struct_get(chunk, TIME_KEY, line);
        chunk.emit_end(line);
    };
    key(chunk, a);
    key(chunk, b);
    cmp(chunk, line);
}

/// The six components `replace` can override, in the order the walker
/// hands them over.
const REPLACE_PROPS: [&str; 6] = ["year", "month", "day", "hour", "minute", "second"];

/// `value.replace(year=…, …)` — a copy with some components overridden.
/// The walker passes every component positionally, `null` where the call
/// omitted it, so a missing one is filled from the receiver here.
/// Type-preserving: a `date`'s replace is a `date`.
/// Stack: `[recv, y, m, d, h, mi, s]` → `[value]`.
pub fn emit_dt_replace(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let base = chunk.alloc_scratch(6);
    let recv = chunk.alloc_scratch(1);
    let tag = chunk.alloc_scratch(1);

    for i in (0..6u16).rev() {
        chunk.emit_op_u16(Op::LOCAL_SET, base + i, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, recv, line);

    chunk.emit_op_u16(Op::LOCAL_GET, recv, line);
    struct_get(chunk, TYPE_KEY, line);
    chunk.emit_op_u16(Op::LOCAL_SET, tag, line);

    for (i, prop) in REPLACE_PROPS.iter().enumerate() {
        chunk.emit_op_u16(Op::LOCAL_GET, base + i as u16, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, recv, line);
        struct_get(chunk, prop, line);
        chunk.emit_op_u16(Op::LOCAL_SET, base + i as u16, line);
        chunk.emit_end(line);
    }

    let slots = [base, base + 1, base + 2, base + 3, base + 4, base + 5];
    emit_utc_from_locals(chunk, &slots, line);
    emit_materialize_tag(chunk, Tag::Local(tag), line);
}

/// `__py_dt_pad(value, width)` — the zero-padded component `strftime`
/// expands to. Stack: `[value, width]` → `[string]`.
pub fn emit_dt_pad(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let w = chunk.alloc_scratch(1);
    let s = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, w, line);
    let to_str = chunk.add_import("ecma:number", "toString");
    chunk.emit_call(to_str, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, s, line);
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_op_u16(Op::LOCAL_GET, w, line);
    chunk.emit_string_const("0", line);
    let pad = chunk.add_import("ecma:string", "padStart");
    chunk.emit_call(pad, 3, line);
}

/// Zero-pad a numeric component to `width`. Stack: `[num]` → `[string]`.
fn emit_pad(chunk: &mut Chunk, width: i32, line: u32) {
    let s = chunk.alloc_scratch(1);
    let to_str = chunk.add_import("ecma:number", "toString");
    chunk.emit_call(to_str, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, s, line);
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    core_wasm::i32_const(chunk, line, width);
    chunk.emit_string_const("0", line);
    let pad = chunk.add_import("ecma:string", "padStart");
    chunk.emit_call(pad, 3, line);
}

/// Append `sep` then the zero-padded `prop` of the value in `obj`.
/// Stack: `[acc]` → `[acc + sep + padded]`.
fn emit_append_part(chunk: &mut Chunk, obj: u16, sep: &str, prop: &str, width: i32, line: u32) {
    if !sep.is_empty() {
        chunk.emit_string_const(sep, line);
        let concat = chunk.add_import("wasm:js-string", "concat");
        chunk.emit_call(concat, 2, line);
    }
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    struct_get(chunk, prop, line);
    emit_pad(chunk, width, line);
    let concat = chunk.add_import("wasm:js-string", "concat");
    chunk.emit_call(concat, 2, line);
}

/// `isoformat()`. Stack: `[value]` → `[string]`.
///
/// One profile entry serves every datetime type (`value_methods` are keyed
/// by name alone), so the receiver's `__type` selects the format:
/// date → `YYYY-MM-DD`, time → `HH:MM:SS`, datetime → both, `T`-joined.
pub fn emit_isoformat(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let obj = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, obj, line);

    emit_is_type(chunk, obj, TYPE_TIME, line);
    chunk.emit_if_value(line);
    emit_time_part(chunk, obj, line);
    chunk.emit_else(line);
    emit_date_part(chunk, obj, line);
    emit_is_type(chunk, obj, TYPE_DATETIME, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("T", line);
    let concat = chunk.add_import("wasm:js-string", "concat");
    chunk.emit_call(concat, 2, line);
    emit_time_part(chunk, obj, line);
    let concat = chunk.add_import("wasm:js-string", "concat");
    chunk.emit_call(concat, 2, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// `obj.__type == tag` as a branch condition. Stack: `[]` → `[i32]`.
fn emit_is_type(chunk: &mut Chunk, obj: u16, tag: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    struct_get(chunk, TYPE_KEY, line);
    chunk.emit_string_const(tag, line);
    let eq = chunk.add_import("wasm:js-string", "equals");
    chunk.emit_call(eq, 2, line);
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
}

/// `HH:MM:SS` from the value in `obj`. Stack: `[]` → `[string]`.
fn emit_time_part(chunk: &mut Chunk, obj: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    struct_get(chunk, "hour", line);
    emit_pad(chunk, 2, line);
    emit_append_part(chunk, obj, ":", "minute", 2, line);
    emit_append_part(chunk, obj, ":", "second", 2, line);
}

/// `YYYY-MM-DD` from the value in `obj`. Stack: `[]` → `[string]`.
fn emit_date_part(chunk: &mut Chunk, obj: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    struct_get(chunk, "year", line);
    emit_pad(chunk, 4, line);
    emit_append_part(chunk, obj, "-", "month", 2, line);
    emit_append_part(chunk, obj, "-", "day", 2, line);
}

/// `date.weekday()` → Monday=0. Stack: `[date]` → `[num]`.
pub fn emit_date_weekday(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    struct_get(chunk, TIME_KEY, line);
    let get_day = chunk.add_import("ecma:date", "getUTCDay");
    chunk.emit_call(get_day, 1, line);
    // ECMA Sunday=0 → Python Monday=0. WASM has no `f64.rem`, and the
    // i32 ops coerce the getter's f64 result, so shift in integer space.
    core_wasm::i32_const(chunk, line, 6);
    chunk.emit_op(Op::I32_ADD, line);
    core_wasm::i32_const(chunk, line, 7);
    chunk.emit_op(Op::I32_REM_S, line);
}
