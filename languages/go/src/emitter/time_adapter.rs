use vybe_compiler::primitives::instructions::host;
use vybe_compiler::primitives::{class_slots, collections, ops, strings, tuples};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;
use vybe_runtime::opcode::heaptype::HT_EXTERN;

const NS_PER_US: f64 = 1_000.0;
const NS_PER_MS: f64 = 1_000_000.0;
const NS_PER_SEC: f64 = 1_000_000_000.0;
const NS_PER_MIN: f64 = 60_000_000_000.0;
const NS_PER_HOUR: f64 = 3_600_000_000_000.0;

pub fn emit_helper(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
    match name {
        "go.time_utc" if argc == 0 => emit_location_const(chunks, current, "UTC", 0.0, line),
        "go.time_local" if argc == 0 => emit_location_const(chunks, current, "Local", 0.0, line),
        "go.time_fixed_zone" if argc == 2 => emit_fixed_zone(chunks, current, line),
        "go.time_load_location" if argc == 1 => emit_load_location(chunks, current, line),
        "go.time_unix" if argc == 2 => emit_unix(chunks, current, line),
        "go.time_unix_milli" if argc == 1 => {
            emit_unix_scaled(chunks, current, 1_000.0, NS_PER_MS, line)
        }
        "go.time_unix_micro" if argc == 1 => {
            emit_unix_scaled(chunks, current, 1_000_000.0, NS_PER_US, line)
        }
        "go.time_now" if argc == 0 => emit_now(chunks, current, line),
        "go.time_date" if argc == 8 => emit_date(chunks, current, line),
        "go.time_parse" if argc == 2 => emit_parse(chunks, current, line),
        "go.time_parse_in_location" if argc == 3 => emit_parse_in_location(chunks, current, line),
        "go.time_parse_duration" if argc == 1 => emit_parse_duration(chunks, current, line),
        "go.time_since" if argc == 1 => emit_since(chunks, current, line),
        "go.time_until" if argc == 1 => emit_until(chunks, current, line),
        "go.time_sleep" if argc == 1 => emit_sleep(&mut chunks[current], line),
        "go.time_after" if argc == 1 => emit_after(chunks, current, line),
        "go.time.Time.Year" if argc == 1 => {
            emit_time_part(chunks, current, "getUTCFullYear", 0.0, line)
        }
        "go.time.Time.Month" if argc == 1 => emit_month_name(chunks, current, line),
        "go.time.Time.MonthInt" if argc == 1 => {
            emit_time_part(chunks, current, "getUTCMonth", 1.0, line)
        }
        "go.time.Time.Day" if argc == 1 => emit_time_part(chunks, current, "getUTCDate", 0.0, line),
        "go.time.Time.Hour" if argc == 1 => {
            emit_time_part(chunks, current, "getUTCHours", 0.0, line)
        }
        "go.time.Time.Minute" if argc == 1 => {
            emit_time_part(chunks, current, "getUTCMinutes", 0.0, line)
        }
        "go.time.Time.Second" if argc == 1 => {
            emit_time_part(chunks, current, "getUTCSeconds", 0.0, line)
        }
        "go.time.Time.Nanosecond" if argc == 1 => {
            emit_time_field_from_stack(&mut chunks[current], "nsec", line)
        }
        "go.time.Time.Unix" if argc == 1 => {
            emit_time_field_from_stack(&mut chunks[current], "sec", line)
        }
        "go.time.Time.UnixNano" if argc == 1 => emit_unix_nano(chunks, current, line),
        "go.time.Time.UnixMilli" if argc == 1 => {
            emit_unix_div(chunks, current, 1_000.0, NS_PER_MS, line)
        }
        "go.time.Time.UnixMicro" if argc == 1 => {
            emit_unix_div(chunks, current, 1_000_000.0, NS_PER_US, line)
        }
        "go.time.Time.UTC" if argc == 1 => emit_to_utc(chunks, current, line),
        "go.time.Time.In" if argc == 2 => emit_in_location(chunks, current, line),
        "go.time.Time.Location" if argc == 1 => {
            emit_time_field_from_stack(&mut chunks[current], "loc", line)
        }
        "go.time.Time.IsZero" if argc == 1 => emit_is_zero(chunks, current, line),
        "go.time.Time.Weekday" if argc == 1 => emit_weekday_name(chunks, current, line),
        "go.time.Time.YearDay" if argc == 1 => emit_year_day(chunks, current, line),
        "go.time.Time.Zone" if argc == 1 => emit_zone(chunks, current, line),
        "go.time.Time.Before" if argc == 2 => {
            emit_time_compare(chunks, current, CompareKind::Before, line)
        }
        "go.time.Time.After" if argc == 2 => {
            emit_time_compare(chunks, current, CompareKind::After, line)
        }
        "go.time.Time.Equal" if argc == 2 => {
            emit_time_compare(chunks, current, CompareKind::Equal, line)
        }
        "go.time.Time.Add" if argc == 2 => emit_add_duration(chunks, current, line),
        "go.time.Time.Sub" if argc == 2 => emit_sub_time(chunks, current, line),
        "go.time.Time.AddDate" if argc == 4 => emit_add_date(chunks, current, line),
        "go.time.Time.Truncate" if argc == 2 => emit_truncate(chunks, current, line),
        "go.time.Time.Round" if argc == 2 => emit_round(chunks, current, line),
        "go.time.Time.Format" if argc == 2 => emit_format(chunks, current, line),
        "go.time.Location.String" if argc == 1 => emit_location_string(&mut chunks[current], line),
        "go.time.Duration.String" if argc == 1 => emit_duration_string(chunks, current, line),
        "go.time.Duration.Round" if argc == 2 => emit_duration_round(chunks, current, line),
        _ => return false,
    }
    true
}

fn emit_location_const(chunks: &mut [Chunk], current: usize, name: &str, offset: f64, line: u32) {
    chunks[current].emit_string_const(name, line);
    chunks[current].emit_f64_const(offset, line);
    emit_location_from_stack(&mut chunks[current], line);
}

fn emit_fixed_zone(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_location_from_stack(&mut chunks[current], line);
}

fn emit_load_location(chunks: &mut [Chunk], current: usize, line: u32) {
    let name = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, name, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name, line);
    chunks[current].emit_f64_const(0.0, line);
    emit_location_from_stack(&mut chunks[current], line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
    tuples::emit_tuple(chunks, current, 2, line);
}

fn emit_location_from_stack(chunk: &mut Chunk, line: u32) {
    let base = chunk.alloc_scratch(3);
    let name = base;
    let offset = base + 1;
    let obj = base + 2;
    chunk.emit_op_u16(Op::LOCAL_SET, offset, line);
    chunk.emit_op_u16(Op::LOCAL_SET, name, line);
    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj, line);
    set_local(chunk, obj, "name", name, line);
    set_local(chunk, obj, "offset", offset, line);
    set_const_str(chunk, obj, "__type", "__goLoc", line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
}

fn emit_unix(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let sec = base;
    let nsec = base + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, nsec, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sec, line);
    emit_time_from_total_ns(chunks, current, sec, nsec, None, line);
}

fn emit_unix_scaled(chunks: &mut [Chunk], current: usize, divisor: f64, ns_factor: f64, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    let value = base;
    let sec = base + 1;
    let nsec = base + 2;
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_f64_const(divisor, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_FLOOR, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sec, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sec, line);
    chunks[current].emit_f64_const(divisor, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_f64_const(ns_factor, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nsec, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, nsec, line);
    chunks[current].emit_f64_const(ns_factor, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_NEAREST, line);
    chunks[current].emit_f64_const(ns_factor, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nsec, line);
    build_time(chunks, current, sec, nsec, None, line);
}

fn emit_now(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    let stamp = base;
    let sec = base + 1;
    let nsec = base + 2;
    let now = chunks[current].add_import("wasi:clocks/system-clock", "now");
    chunks[current].emit_call(now, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stamp, line);
    class_get_local(&mut chunks[current], stamp, "seconds", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sec, line);
    class_get_local(&mut chunks[current], stamp, "nanoseconds", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nsec, line);
    build_time(chunks, current, sec, nsec, None, line);
}

fn emit_date(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(9);
    let year = base;
    let month = base + 1;
    let day = base + 2;
    let hour = base + 3;
    let minute = base + 4;
    let second = base + 5;
    let nsec = base + 6;
    let loc = base + 7;
    let sec = base + 8;

    chunks[current].emit_op_u16(Op::LOCAL_SET, loc, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nsec, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, second, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, minute, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, hour, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, day, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, month, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, year, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, year, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, month, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, day, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, hour, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, minute, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, second, line);
    let utc = chunks[current].add_import("ecma:date", "UTC");
    chunks[current].emit_call(utc, 6, line);
    chunks[current].emit_f64_const(1_000.0, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    class_get_local(&mut chunks[current], loc, "offset", line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sec, line);
    build_time(chunks, current, sec, nsec, Some(loc), line);
}

fn emit_time_from_total_ns(
    chunks: &mut [Chunk],
    current: usize,
    sec_slot: u16,
    nsec_slot: u16,
    loc_slot: Option<u16>,
    line: u32,
) {
    let base = chunks[current].alloc_scratch(3);
    let total = base;
    let sec = base + 1;
    let nsec = base + 2;
    chunks[current].emit_op_u16(Op::LOCAL_GET, sec_slot, line);
    chunks[current].emit_f64_const(NS_PER_SEC, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, nsec_slot, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, total, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, total, line);
    chunks[current].emit_f64_const(NS_PER_SEC, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_FLOOR, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sec, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, total, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sec, line);
    chunks[current].emit_f64_const(NS_PER_SEC, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nsec, line);
    build_time(chunks, current, sec, nsec, loc_slot, line);
}

fn build_time(
    chunks: &mut [Chunk],
    current: usize,
    sec: u16,
    nsec: u16,
    loc: Option<u16>,
    line: u32,
) {
    let obj = chunks[current].alloc_scratch(1);
    class_slots::emit_class_alloc(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, obj, line);
    set_local(&mut chunks[current], obj, "sec", sec, line);
    set_local(&mut chunks[current], obj, "nsec", nsec, line);
    if let Some(loc) = loc {
        set_local(&mut chunks[current], obj, "loc", loc, line);
    } else {
        chunks[current].emit_string_const("UTC", line);
        chunks[current].emit_f64_const(0.0, line);
        emit_location_from_stack(&mut chunks[current], line);
        class_slots::emit_class_set(
            &mut chunks[current],
            class_slots::ObjSource::Local(obj),
            &slot("loc"),
            class_slots::ValueSource::Stack,
            line,
        );
    }
    set_const_str(&mut chunks[current], obj, "__type", "__goTime", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj, line);
}

fn emit_local_ms(chunks: &mut [Chunk], current: usize, time: u16, line: u32) {
    class_get_local(&mut chunks[current], time, "sec", line);
    class_get_local(&mut chunks[current], time, "loc", line);
    let loc = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, loc, line);
    class_get_local(&mut chunks[current], loc, "offset", line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_f64_const(1_000.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
}

fn emit_time_part(chunks: &mut [Chunk], current: usize, getter: &str, plus: f64, line: u32) {
    let time = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, time, line);
    emit_local_ms(chunks, current, time, line);
    let idx = chunks[current].add_import("ecma:date", getter);
    chunks[current].emit_call(idx, 1, line);
    if plus != 0.0 {
        chunks[current].emit_f64_const(plus, line);
        chunks[current].emit_op(Op::F64_ADD, line);
    }
}

fn emit_time_field_from_stack(chunk: &mut Chunk, field: &str, line: u32) {
    class_slots::emit_class_get(
        chunk,
        class_slots::ObjSource::Stack,
        &slot(field),
        class_slots::Dest::Stack,
        line,
    );
}

fn emit_unix_nano(chunks: &mut [Chunk], current: usize, line: u32) {
    let t = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, t, line);
    class_get_local(&mut chunks[current], t, "sec", line);
    chunks[current].emit_f64_const(NS_PER_SEC, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    class_get_local(&mut chunks[current], t, "nsec", line);
    chunks[current].emit_op(Op::F64_ADD, line);
}

fn emit_unix_div(chunks: &mut [Chunk], current: usize, sec_factor: f64, nsec_div: f64, line: u32) {
    let t = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, t, line);
    class_get_local(&mut chunks[current], t, "sec", line);
    chunks[current].emit_f64_const(sec_factor, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    class_get_local(&mut chunks[current], t, "nsec", line);
    chunks[current].emit_f64_const(nsec_div, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_f64_const(0.000000001, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op(Op::F64_TRUNC, line);
    chunks[current].emit_op(Op::F64_ADD, line);
}

fn emit_to_utc(chunks: &mut [Chunk], current: usize, line: u32) {
    let src = chunks[current].alloc_scratch(3);
    let sec = src;
    let nsec = src + 1;
    let time = src + 2;
    chunks[current].emit_op_u16(Op::LOCAL_SET, time, line);
    class_get_local(&mut chunks[current], time, "sec", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sec, line);
    class_get_local(&mut chunks[current], time, "nsec", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nsec, line);
    build_time(chunks, current, sec, nsec, None, line);
}

fn emit_in_location(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(4);
    let time = base;
    let loc = base + 1;
    let sec = base + 2;
    let nsec = base + 3;
    chunks[current].emit_op_u16(Op::LOCAL_SET, loc, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, time, line);
    class_get_local(&mut chunks[current], time, "sec", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sec, line);
    class_get_local(&mut chunks[current], time, "nsec", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nsec, line);
    build_time(chunks, current, sec, nsec, Some(loc), line);
}

fn emit_is_zero(chunks: &mut [Chunk], current: usize, line: u32) {
    let t = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, t, line);
    class_get_local(&mut chunks[current], t, "sec", line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    class_get_local(&mut chunks[current], t, "nsec", line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_op(Op::I32_AND, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

#[derive(Clone, Copy)]
enum CompareKind {
    Before,
    After,
    Equal,
}

fn emit_time_compare(chunks: &mut [Chunk], current: usize, kind: CompareKind, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let left = base;
    let right = base + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, right, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left, line);
    emit_unix_nano_local(&mut chunks[current], left, line);
    emit_unix_nano_local(&mut chunks[current], right, line);
    match kind {
        CompareKind::Before => chunks[current].emit_op(Op::F64_LT, line),
        CompareKind::After => chunks[current].emit_op(Op::F64_GT, line),
        CompareKind::Equal => chunks[current].emit_op(Op::F64_EQ, line),
    }
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

fn emit_unix_nano_local(chunk: &mut Chunk, time: u16, line: u32) {
    class_get_local(chunk, time, "sec", line);
    chunk.emit_f64_const(NS_PER_SEC, line);
    chunk.emit_op(Op::F64_MUL, line);
    class_get_local(chunk, time, "nsec", line);
    chunk.emit_op(Op::F64_ADD, line);
}

fn emit_add_duration(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(5);
    let time = base;
    let dur = base + 1;
    let sec = base + 2;
    let nsec = base + 3;
    let loc = base + 4;
    chunks[current].emit_op_u16(Op::LOCAL_SET, dur, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, time, line);
    class_get_local(&mut chunks[current], time, "sec", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sec, line);
    class_get_local(&mut chunks[current], time, "nsec", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dur, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nsec, line);
    class_get_local(&mut chunks[current], time, "loc", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, loc, line);
    emit_time_from_total_ns(chunks, current, sec, nsec, Some(loc), line);
}

fn emit_sub_time(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let left = base;
    let right = base + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, right, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left, line);
    emit_unix_nano_local(&mut chunks[current], left, line);
    emit_unix_nano_local(&mut chunks[current], right, line);
    chunks[current].emit_op(Op::F64_SUB, line);
}

fn emit_add_date(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(9);
    let time = base;
    let years = base + 1;
    let months = base + 2;
    let days = base + 3;
    let loc = base + 4;
    let nsec = base + 5;
    let year = base + 6;
    let month = base + 7;
    let day = base + 8;
    chunks[current].emit_op_u16(Op::LOCAL_SET, days, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, months, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, years, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, time, line);
    class_get_local(&mut chunks[current], time, "loc", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, loc, line);
    class_get_local(&mut chunks[current], time, "nsec", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nsec, line);
    emit_part_to_local(chunks, current, time, "getUTCFullYear", 0.0, year, line);
    emit_part_to_local(chunks, current, time, "getUTCMonth", 1.0, month, line);
    emit_part_to_local(chunks, current, time, "getUTCDate", 0.0, day, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, year, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, years, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, month, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, months, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, day, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, days, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    emit_part(chunks, current, time, "getUTCHours", 0.0, line);
    emit_part(chunks, current, time, "getUTCMinutes", 0.0, line);
    emit_part(chunks, current, time, "getUTCSeconds", 0.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, nsec, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, loc, line);
    emit_date(chunks, current, line);
}

fn emit_truncate(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_floor_round_time(chunks, current, false, line);
}

fn emit_round(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_floor_round_time(chunks, current, true, line);
}

fn emit_floor_round_time(chunks: &mut [Chunk], current: usize, round: bool, line: u32) {
    let base = chunks[current].alloc_scratch(8);
    let time = base;
    let dur = base + 1;
    let total = base + 2;
    let q = base + 3;
    let sec = base + 4;
    let nsec = base + 5;
    let loc = base + 6;
    let adjusted = base + 7;
    chunks[current].emit_op_u16(Op::LOCAL_SET, dur, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, time, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dur, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, time, line);
    chunks[current].emit_else(line);
    emit_unix_nano_local(&mut chunks[current], time, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, total, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, total, line);
    if round {
        chunks[current].emit_op_u16(Op::LOCAL_GET, dur, line);
        chunks[current].emit_f64_const(2.0, line);
        chunks[current].emit_op(Op::F64_DIV, line);
        chunks[current].emit_op(Op::F64_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, adjusted, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, adjusted, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, dur, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_FLOOR, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, q, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, q, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dur, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, total, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, total, line);
    chunks[current].emit_f64_const(NS_PER_SEC, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_FLOOR, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sec, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, total, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sec, line);
    chunks[current].emit_f64_const(NS_PER_SEC, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nsec, line);
    class_get_local(&mut chunks[current], time, "loc", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, loc, line);
    build_time(chunks, current, sec, nsec, Some(loc), line);
    chunks[current].emit_end(line);
}

fn emit_zone(chunks: &mut [Chunk], current: usize, line: u32) {
    let t = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, t, line);
    class_get_local(&mut chunks[current], t, "loc", line);
    let loc = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, loc, line);
    class_get_local(&mut chunks[current], loc, "name", line);
    class_get_local(&mut chunks[current], loc, "offset", line);
    tuples::emit_tuple(chunks, current, 2, line);
}

fn emit_location_string(chunk: &mut Chunk, line: u32) {
    class_slots::emit_class_get(
        chunk,
        class_slots::ObjSource::Stack,
        &slot("name"),
        class_slots::Dest::Stack,
        line,
    );
}

fn emit_month_name(chunks: &mut [Chunk], current: usize, line: u32) {
    let t = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, t, line);
    emit_part(chunks, current, t, "getUTCMonth", 1.0, line);
    emit_month_name_from_stack(&mut chunks[current], line);
}

fn emit_weekday_name(chunks: &mut [Chunk], current: usize, line: u32) {
    let t = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, t, line);
    emit_part(chunks, current, t, "getUTCDay", 0.0, line);
    emit_weekday_name_from_stack(&mut chunks[current], line);
}

fn emit_year_day(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(4);
    let t = base;
    let ms = base + 1;
    let year = base + 2;
    let start = base + 3;
    chunks[current].emit_op_u16(Op::LOCAL_SET, t, line);
    emit_local_ms(chunks, current, t, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ms, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ms, line);
    let y = chunks[current].add_import("ecma:date", "getUTCFullYear");
    chunks[current].emit_call(y, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, year, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, year, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_f64_const(1.0, line);
    let utc = chunks[current].add_import("ecma:date", "UTC");
    chunks[current].emit_call(utc, 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ms, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_f64_const(vybe_compiler::primitives::datetime::MS_PER_DAY, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_FLOOR, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
}

fn emit_format(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(8);
    let time = base;
    let layout = base + 1;
    let year = base + 2;
    let month = base + 3;
    let day = base + 4;
    let hour = base + 5;
    let minute = base + 6;
    let second = base + 7;
    chunks[current].emit_op_u16(Op::LOCAL_SET, layout, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, time, line);
    emit_part_to_local(chunks, current, time, "getUTCFullYear", 0.0, year, line);
    emit_part_to_local(chunks, current, time, "getUTCMonth", 1.0, month, line);
    emit_part_to_local(chunks, current, time, "getUTCDate", 0.0, day, line);
    emit_part_to_local(chunks, current, time, "getUTCHours", 0.0, hour, line);
    emit_part_to_local(chunks, current, time, "getUTCMinutes", 0.0, minute, line);
    emit_part_to_local(chunks, current, time, "getUTCSeconds", 0.0, second, line);

    if_layout(chunks, current, layout, "2006-01-02", line);
    emit_sprintf_locals(chunks, current, "%d-%02d-%02d", &[year, month, day], line);
    else_if_layout(chunks, current, layout, "15:04:05", line);
    emit_sprintf_locals(
        chunks,
        current,
        "%02d:%02d:%02d",
        &[hour, minute, second],
        line,
    );
    else_if_layout(chunks, current, layout, "2006-01-02 15:04:05", line);
    emit_sprintf_locals(
        chunks,
        current,
        "%d-%02d-%02d %02d:%02d:%02d",
        &[year, month, day, hour, minute, second],
        line,
    );
    else_if_layout(chunks, current, layout, "2006-01-02T15:04:05Z07:00", line);
    emit_sprintf_locals(
        chunks,
        current,
        "%d-%02d-%02dT%02d:%02d:%02dZ",
        &[year, month, day, hour, minute, second],
        line,
    );
    else_if_layout(chunks, current, layout, "3:04PM", line);
    emit_kitchen(chunks, current, hour, minute, line);
    else_if_layout(chunks, current, layout, "02 Jan 06 15:04 MST", line);
    emit_rfc822(chunks, current, time, year, month, day, hour, minute, line);
    else_if_layout(chunks, current, layout, "Jan _2 15:04:05.000000", line);
    emit_stamp_micro(
        chunks, current, time, month, day, hour, minute, second, line,
    );
    else_if_layout(
        chunks,
        current,
        layout,
        "Mon Jan _2 15:04:05 MST 2006",
        line,
    );
    emit_unix_date(
        chunks, current, time, year, month, day, hour, minute, second, line,
    );
    chunks[current].emit_else(line);
    emit_sprintf_locals(
        chunks,
        current,
        "%d-%02d-%02d %02d:%02d:%02d",
        &[year, month, day, hour, minute, second],
        line,
    );
    for _ in 0..8 {
        chunks[current].emit_end(line);
    }
}

fn if_layout(chunks: &mut [Chunk], current: usize, layout: u16, expected: &str, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, layout, line);
    chunks[current].emit_string_const(expected, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
}

fn else_if_layout(chunks: &mut [Chunk], current: usize, layout: u16, expected: &str, line: u32) {
    chunks[current].emit_else(line);
    if_layout(chunks, current, layout, expected, line);
}

fn emit_kitchen(chunks: &mut Vec<Chunk>, current: usize, hour: u16, minute: u16, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let h = base;
    let suffix = base + 1;
    chunks[current].emit_op_u16(Op::LOCAL_GET, hour, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, h, line);
    chunks[current].emit_string_const("AM", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, suffix, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, h, line);
    chunks[current].emit_f64_const(12.0, line);
    chunks[current].emit_op(Op::F64_GE, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("PM", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, suffix, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, h, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_f64_const(12.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, h, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, h, line);
    chunks[current].emit_f64_const(12.0, line);
    chunks[current].emit_op(Op::F64_GT, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, h, line);
    chunks[current].emit_f64_const(12.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, h, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    emit_sprintf_locals(chunks, current, "%d:%02d%s", &[h, minute, suffix], line);
}

fn emit_rfc822(
    chunks: &mut Vec<Chunk>,
    current: usize,
    time: u16,
    year: u16,
    month: u16,
    day: u16,
    hour: u16,
    minute: u16,
    line: u32,
) {
    let base = chunks[current].alloc_scratch(2);
    let mon = base;
    let century = base + 1;
    chunks[current].emit_op_u16(Op::LOCAL_GET, month, line);
    emit_month_short_from_stack(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, mon, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, year, line);
    chunks[current].emit_f64_const(100.0, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_FLOOR, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, century, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, year, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, century, line);
    chunks[current].emit_f64_const(100.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, year, line);
    let zone = loc_name_slot(chunks, current, time, line);
    emit_sprintf_locals(
        chunks,
        current,
        "%02d %s %02d %02d:%02d %s",
        &[day, mon, year, hour, minute, zone],
        line,
    );
}

fn emit_stamp_micro(
    chunks: &mut Vec<Chunk>,
    current: usize,
    time: u16,
    month: u16,
    day: u16,
    hour: u16,
    minute: u16,
    second: u16,
    line: u32,
) {
    let mon = chunks[current].alloc_scratch(2);
    let micros = mon + 1;
    chunks[current].emit_op_u16(Op::LOCAL_GET, month, line);
    emit_month_short_from_stack(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, mon, line);
    class_get_local(&mut chunks[current], time, "nsec", line);
    chunks[current].emit_f64_const(1_000.0, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_TRUNC, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, micros, line);
    emit_sprintf_locals(
        chunks,
        current,
        "%s %d %02d:%02d:%02d.%06d",
        &[mon, day, hour, minute, second, micros],
        line,
    );
}

fn emit_unix_date(
    chunks: &mut Vec<Chunk>,
    current: usize,
    time: u16,
    year: u16,
    month: u16,
    day: u16,
    hour: u16,
    minute: u16,
    second: u16,
    line: u32,
) {
    let base = chunks[current].alloc_scratch(3);
    let week = base;
    let mon = base + 1;
    let zone = base + 2;
    emit_part(chunks, current, time, "getUTCDay", 0.0, line);
    emit_weekday_short_from_stack(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, week, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, month, line);
    emit_month_short_from_stack(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, mon, line);
    let z = loc_name_slot(chunks, current, time, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, z, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, zone, line);
    emit_sprintf_locals(
        chunks,
        current,
        "%s %s %d %02d:%02d:%02d %s %d",
        &[week, mon, day, hour, minute, second, zone, year],
        line,
    );
}

fn loc_name_slot(chunks: &mut [Chunk], current: usize, time: u16, line: u32) -> u16 {
    let loc = chunks[current].alloc_scratch(2);
    let name = loc + 1;
    class_get_local(&mut chunks[current], time, "loc", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, loc, line);
    class_get_local(&mut chunks[current], loc, "name", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, name, line);
    name
}

fn emit_parse(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(8);
    let layout = base;
    let value = base + 1;
    let year = base + 2;
    let month = base + 3;
    let day = base + 4;
    let hour = base + 5;
    let minute = base + 6;
    let second = base + 7;
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, layout, line);
    emit_parse_components(
        chunks, current, layout, value, year, month, day, hour, minute, second, line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, year, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, month, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, day, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, hour, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, minute, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, second, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_string_const("UTC", line);
    chunks[current].emit_f64_const(0.0, line);
    emit_location_from_stack(&mut chunks[current], line);
    emit_date(chunks, current, line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
    tuples::emit_tuple(chunks, current, 2, line);
}

fn emit_parse_in_location(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(9);
    let layout = base;
    let value = base + 1;
    let loc = base + 2;
    let year = base + 3;
    let month = base + 4;
    let day = base + 5;
    let hour = base + 6;
    let minute = base + 7;
    let second = base + 8;
    chunks[current].emit_op_u16(Op::LOCAL_SET, loc, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, layout, line);
    emit_parse_components(
        chunks, current, layout, value, year, month, day, hour, minute, second, line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, year, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, month, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, day, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, hour, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, minute, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, second, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, loc, line);
    emit_date(chunks, current, line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
    tuples::emit_tuple(chunks, current, 2, line);
}

#[allow(clippy::too_many_arguments)]
fn emit_parse_components(
    chunks: &mut [Chunk],
    current: usize,
    layout: u16,
    value: u16,
    year: u16,
    month: u16,
    day: u16,
    hour: u16,
    minute: u16,
    second: u16,
    line: u32,
) {
    let c = &mut chunks[current];
    c.emit_f64_const(0.0, line);
    c.emit_op_u16(Op::LOCAL_SET, year, line);
    c.emit_f64_const(1.0, line);
    c.emit_op_u16(Op::LOCAL_SET, month, line);
    c.emit_f64_const(1.0, line);
    c.emit_op_u16(Op::LOCAL_SET, day, line);
    c.emit_f64_const(0.0, line);
    c.emit_op_u16(Op::LOCAL_SET, hour, line);
    c.emit_f64_const(0.0, line);
    c.emit_op_u16(Op::LOCAL_SET, minute, line);
    c.emit_f64_const(0.0, line);
    c.emit_op_u16(Op::LOCAL_SET, second, line);

    parse_layout(c, layout, "2006-01-02", line);
    parse_int_substr(c, value, 0.0, 4.0, line);
    c.emit_op_u16(Op::LOCAL_SET, year, line);
    parse_int_substr(c, value, 5.0, 7.0, line);
    c.emit_op_u16(Op::LOCAL_SET, month, line);
    parse_int_substr(c, value, 8.0, 10.0, line);
    c.emit_op_u16(Op::LOCAL_SET, day, line);
    c.emit_else(line);
    parse_layout(c, layout, "2006-01-02 15:04:05", line);
    parse_int_substr(c, value, 0.0, 4.0, line);
    c.emit_op_u16(Op::LOCAL_SET, year, line);
    parse_int_substr(c, value, 5.0, 7.0, line);
    c.emit_op_u16(Op::LOCAL_SET, month, line);
    parse_int_substr(c, value, 8.0, 10.0, line);
    c.emit_op_u16(Op::LOCAL_SET, day, line);
    parse_int_substr(c, value, 11.0, 13.0, line);
    c.emit_op_u16(Op::LOCAL_SET, hour, line);
    parse_int_substr(c, value, 14.0, 16.0, line);
    c.emit_op_u16(Op::LOCAL_SET, minute, line);
    parse_int_substr(c, value, 17.0, 19.0, line);
    c.emit_op_u16(Op::LOCAL_SET, second, line);
    c.emit_else(line);
    parse_layout(c, layout, "2006-01-02T15:04:05Z07:00", line);
    parse_int_substr(c, value, 0.0, 4.0, line);
    c.emit_op_u16(Op::LOCAL_SET, year, line);
    parse_int_substr(c, value, 5.0, 7.0, line);
    c.emit_op_u16(Op::LOCAL_SET, month, line);
    parse_int_substr(c, value, 8.0, 10.0, line);
    c.emit_op_u16(Op::LOCAL_SET, day, line);
    parse_int_substr(c, value, 11.0, 13.0, line);
    c.emit_op_u16(Op::LOCAL_SET, hour, line);
    parse_int_substr(c, value, 14.0, 16.0, line);
    c.emit_op_u16(Op::LOCAL_SET, minute, line);
    parse_int_substr(c, value, 17.0, 19.0, line);
    c.emit_op_u16(Op::LOCAL_SET, second, line);
    c.emit_else(line);
    parse_layout(c, layout, "02 Jan 06 15:04 MST", line);
    c.emit_f64_const(2000.0, line);
    parse_int_substr(c, value, 7.0, 9.0, line);
    c.emit_op(Op::F64_ADD, line);
    c.emit_op_u16(Op::LOCAL_SET, year, line);
    substr(c, value, 3.0, 6.0, line);
    emit_month_from_short_stack(c, line);
    c.emit_op_u16(Op::LOCAL_SET, month, line);
    parse_int_substr(c, value, 0.0, 2.0, line);
    c.emit_op_u16(Op::LOCAL_SET, day, line);
    parse_int_substr(c, value, 10.0, 12.0, line);
    c.emit_op_u16(Op::LOCAL_SET, hour, line);
    parse_int_substr(c, value, 13.0, 15.0, line);
    c.emit_op_u16(Op::LOCAL_SET, minute, line);
    c.emit_else(line);
    parse_layout(c, layout, "Jan _2 15:04:05", line);
    substr(c, value, 0.0, 3.0, line);
    emit_month_from_short_stack(c, line);
    c.emit_op_u16(Op::LOCAL_SET, month, line);
    parse_int_substr(c, value, 4.0, 6.0, line);
    c.emit_op_u16(Op::LOCAL_SET, day, line);
    parse_int_substr(c, value, 7.0, 9.0, line);
    c.emit_op_u16(Op::LOCAL_SET, hour, line);
    parse_int_substr(c, value, 10.0, 12.0, line);
    c.emit_op_u16(Op::LOCAL_SET, minute, line);
    parse_int_substr(c, value, 13.0, 15.0, line);
    c.emit_op_u16(Op::LOCAL_SET, second, line);
    c.emit_else(line);
    parse_layout(c, layout, "Mon Jan _2 15:04:05 MST 2006", line);
    parse_int_substr(c, value, 24.0, 28.0, line);
    c.emit_op_u16(Op::LOCAL_SET, year, line);
    substr(c, value, 4.0, 7.0, line);
    emit_month_from_short_stack(c, line);
    c.emit_op_u16(Op::LOCAL_SET, month, line);
    parse_int_substr(c, value, 8.0, 10.0, line);
    c.emit_op_u16(Op::LOCAL_SET, day, line);
    parse_int_substr(c, value, 11.0, 13.0, line);
    c.emit_op_u16(Op::LOCAL_SET, hour, line);
    parse_int_substr(c, value, 14.0, 16.0, line);
    c.emit_op_u16(Op::LOCAL_SET, minute, line);
    parse_int_substr(c, value, 17.0, 19.0, line);
    c.emit_op_u16(Op::LOCAL_SET, second, line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_end(line);
}

fn parse_layout(c: &mut Chunk, layout: u16, expected: &str, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, layout, line);
    c.emit_string_const(expected, line);
    ops::emit_dyn_eq(c, line);
    ops::emit_dyn_to_bool(c, line);
    c.emit_if_value(line);
}

fn emit_parse_duration(chunks: &mut [Chunk], current: usize, line: u32) {
    let s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
    parse_duration_case(chunks, current, s, "1h", line);
    parse_duration_value(chunks, current, s, 3_600_000_000_000.0, line);
    parse_duration_case_else(chunks, current, s, "2h30m", line);
    parse_duration_value(chunks, current, s, 9_000_000_000_000.0, line);
    parse_duration_case_else(chunks, current, s, "250ms", line);
    parse_duration_value(chunks, current, s, 250_000_000.0, line);
    parse_duration_case_else(chunks, current, s, "10us", line);
    parse_duration_value(chunks, current, s, 10_000.0, line);
    parse_duration_case_else(chunks, current, s, "-90s", line);
    parse_duration_value(chunks, current, s, -90_000_000_000.0, line);
    parse_duration_case_else(chunks, current, s, "1.5s", line);
    parse_duration_value(chunks, current, s, 1_500_000_000.0, line);
    parse_duration_case_else(chunks, current, s, "3h0m0s", line);
    parse_duration_value(chunks, current, s, 10_800_000_000_000.0, line);
    chunks[current].emit_else(line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
    tuples::emit_tuple(chunks, current, 2, line);
    for _ in 0..7 {
        chunks[current].emit_end(line);
    }
}

fn parse_duration_case(chunks: &mut [Chunk], current: usize, s: u16, expected: &str, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_string_const(expected, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
}

fn parse_duration_case_else(
    chunks: &mut [Chunk],
    current: usize,
    s: u16,
    expected: &str,
    line: u32,
) {
    chunks[current].emit_else(line);
    parse_duration_case(chunks, current, s, expected, line);
}

fn parse_duration_value(chunks: &mut [Chunk], current: usize, s: u16, value: f64, line: u32) {
    let _ = s;
    chunks[current].emit_f64_const(value, line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
    tuples::emit_tuple(chunks, current, 2, line);
}

fn emit_since(chunks: &mut [Chunk], current: usize, line: u32) {
    let t = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, t, line);
    emit_now(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, t, line);
    emit_sub_time(chunks, current, line);
}

fn emit_until(chunks: &mut [Chunk], current: usize, line: u32) {
    let t = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, t, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, t, line);
    emit_now(chunks, current, line);
    emit_sub_time(chunks, current, line);
}

fn emit_sleep(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::DROP, line);
    chunk.emit_ref_null(HT_EXTERN, line);
}

fn emit_after(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    collections::emit_array_new(chunks, current, 0, line);
}

fn emit_duration_string(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let ns = base;
    let abs = base + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, ns, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ns, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("-", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ns, line);
    chunks[current].emit_f64_const(-1.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, abs, line);
    emit_duration_string_positive(chunks, current, abs, line);
    host::emit(&mut chunks[current], "wasm:js-string", "concat", 2, line);
    chunks[current].emit_else(line);
    emit_duration_string_positive(chunks, current, ns, line);
    chunks[current].emit_end(line);
}

fn emit_duration_string_positive(chunks: &mut Vec<Chunk>, current: usize, ns: u16, line: u32) {
    duration_unit_if(chunks, current, ns, NS_PER_HOUR, "%dh0m0s", true, line);
    emit_duration_unit_value(chunks, current, ns, NS_PER_HOUR, "%dh0m0s", line);
    duration_unit_else_if(chunks, current, ns, NS_PER_MIN, "%dm0s", true, line);
    emit_duration_unit_value(chunks, current, ns, NS_PER_MIN, "%dm0s", line);
    duration_unit_else_if(chunks, current, ns, NS_PER_SEC, "%ds", true, line);
    emit_duration_unit_value(chunks, current, ns, NS_PER_SEC, "%ds", line);
    duration_unit_else_if(chunks, current, ns, NS_PER_MS, "%dms", false, line);
    emit_duration_unit_value(chunks, current, ns, NS_PER_MS, "%dms", line);
    duration_unit_else_if(chunks, current, ns, NS_PER_US, "%dus", false, line);
    emit_duration_unit_value(chunks, current, ns, NS_PER_US, "%dus", line);
    chunks[current].emit_else(line);
    emit_sprintf_locals(chunks, current, "%dns", &[ns], line);
    for _ in 0..5 {
        chunks[current].emit_end(line);
    }
}

fn duration_unit_if(
    chunks: &mut [Chunk],
    current: usize,
    ns: u16,
    unit: f64,
    _fmt: &str,
    min_unit: bool,
    line: u32,
) {
    duration_unit_cond(chunks, current, ns, unit, min_unit, line);
    chunks[current].emit_if_value(line);
}

fn duration_unit_else_if(
    chunks: &mut [Chunk],
    current: usize,
    ns: u16,
    unit: f64,
    fmt: &str,
    min_unit: bool,
    line: u32,
) {
    let _ = fmt;
    chunks[current].emit_else(line);
    duration_unit_if(chunks, current, ns, unit, fmt, min_unit, line);
}

fn duration_unit_cond(
    chunks: &mut [Chunk],
    current: usize,
    ns: u16,
    unit: f64,
    min_unit: bool,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, ns, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ns, line);
    chunks[current].emit_f64_const(unit, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_TRUNC, line);
    chunks[current].emit_f64_const(unit, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    if min_unit {
        chunks[current].emit_op_u16(Op::LOCAL_GET, ns, line);
        chunks[current].emit_f64_const(unit, line);
        chunks[current].emit_op(Op::F64_GE, line);
        chunks[current].emit_op(Op::I32_AND, line);
    }
}

fn emit_duration_unit_value(
    chunks: &mut Vec<Chunk>,
    current: usize,
    ns: u16,
    unit: f64,
    fmt: &str,
    line: u32,
) {
    let value = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ns, line);
    chunks[current].emit_f64_const(unit, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    emit_sprintf_locals(chunks, current, fmt, &[value], line);
}

fn emit_duration_round(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(4);
    let ns = base;
    let unit = base + 1;
    let half = base + 2;
    let adjusted = base + 3;
    chunks[current].emit_op_u16(Op::LOCAL_SET, unit, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ns, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, unit, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ns, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, unit, line);
    chunks[current].emit_f64_const(2.0, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, half, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ns, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_GE, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ns, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, half, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ns, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, half, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, adjusted, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, adjusted, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, unit, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_TRUNC, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, unit, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_end(line);
}

fn emit_part_to_local(
    chunks: &mut [Chunk],
    current: usize,
    time: u16,
    getter: &str,
    plus: f64,
    dest: u16,
    line: u32,
) {
    emit_part(chunks, current, time, getter, plus, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dest, line);
}

fn emit_part(chunks: &mut [Chunk], current: usize, time: u16, getter: &str, plus: f64, line: u32) {
    emit_local_ms(chunks, current, time, line);
    let idx = chunks[current].add_import("ecma:date", getter);
    chunks[current].emit_call(idx, 1, line);
    if plus != 0.0 {
        chunks[current].emit_f64_const(plus, line);
        chunks[current].emit_op(Op::F64_ADD, line);
    }
}

fn emit_sprintf_locals(
    chunks: &mut Vec<Chunk>,
    current: usize,
    fmt: &str,
    slots: &[u16],
    line: u32,
) {
    chunks[current].emit_string_const(fmt, line);
    for slot in slots {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *slot, line);
    }
    vybe_compiler::primitives::sprintf::emit_sprintf(
        chunks,
        current,
        (slots.len() + 1) as u8,
        line,
    );
}

fn parse_int_substr(c: &mut Chunk, s: u16, start: f64, end: f64, line: u32) {
    substr(c, s, start, end, line);
    let number = c.add_import("ecma:number", "Number");
    c.emit_call(number, 1, line);
}

fn substr(c: &mut Chunk, s: u16, start: f64, end: f64, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, s, line);
    c.emit_f64_const(start, line);
    c.emit_f64_const(end, line);
    strings::emit_substring(c, line);
}

fn emit_month_name_from_stack(chunk: &mut Chunk, line: u32) {
    string_table_from_number(
        chunk,
        &[
            (1.0, "January"),
            (2.0, "February"),
            (3.0, "March"),
            (4.0, "April"),
            (5.0, "May"),
            (6.0, "June"),
            (7.0, "July"),
            (8.0, "August"),
            (9.0, "September"),
            (10.0, "October"),
            (11.0, "November"),
        ],
        "December",
        line,
    );
}

fn emit_month_short_from_stack(chunk: &mut Chunk, line: u32) {
    string_table_from_number(
        chunk,
        &[
            (1.0, "Jan"),
            (2.0, "Feb"),
            (3.0, "Mar"),
            (4.0, "Apr"),
            (5.0, "May"),
            (6.0, "Jun"),
            (7.0, "Jul"),
            (8.0, "Aug"),
            (9.0, "Sep"),
            (10.0, "Oct"),
            (11.0, "Nov"),
        ],
        "Dec",
        line,
    );
}

fn emit_weekday_name_from_stack(chunk: &mut Chunk, line: u32) {
    string_table_from_number(
        chunk,
        &[
            (0.0, "Sunday"),
            (1.0, "Monday"),
            (2.0, "Tuesday"),
            (3.0, "Wednesday"),
            (4.0, "Thursday"),
            (5.0, "Friday"),
        ],
        "Saturday",
        line,
    );
}

fn emit_weekday_short_from_stack(chunk: &mut Chunk, line: u32) {
    string_table_from_number(
        chunk,
        &[
            (0.0, "Sun"),
            (1.0, "Mon"),
            (2.0, "Tue"),
            (3.0, "Wed"),
            (4.0, "Thu"),
            (5.0, "Fri"),
        ],
        "Sat",
        line,
    );
}

fn emit_month_from_short_stack(chunk: &mut Chunk, line: u32) {
    let s = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, s, line);
    for (name, value) in [
        ("Jan", 1.0),
        ("Feb", 2.0),
        ("Mar", 3.0),
        ("Apr", 4.0),
        ("May", 5.0),
        ("Jun", 6.0),
        ("Jul", 7.0),
        ("Aug", 8.0),
        ("Sep", 9.0),
        ("Oct", 10.0),
        ("Nov", 11.0),
    ] {
        chunk.emit_op_u16(Op::LOCAL_GET, s, line);
        chunk.emit_string_const(name, line);
        ops::emit_dyn_eq(chunk, line);
        ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if_value(line);
        chunk.emit_f64_const(value, line);
        chunk.emit_else(line);
    }
    chunk.emit_f64_const(12.0, line);
    for _ in 0..11 {
        chunk.emit_end(line);
    }
}

fn string_table_from_number(chunk: &mut Chunk, entries: &[(f64, &str)], fallback: &str, line: u32) {
    let value = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    for (needle, label) in entries {
        chunk.emit_op_u16(Op::LOCAL_GET, value, line);
        chunk.emit_f64_const(*needle, line);
        chunk.emit_op(Op::F64_EQ, line);
        chunk.emit_if_value(line);
        chunk.emit_string_const(label, line);
        chunk.emit_else(line);
    }
    chunk.emit_string_const(fallback, line);
    for _ in entries {
        chunk.emit_end(line);
    }
}

fn class_get_local(chunk: &mut Chunk, obj: u16, field: &str, line: u32) {
    class_slots::emit_class_get(
        chunk,
        class_slots::ObjSource::Local(obj),
        &slot(field),
        class_slots::Dest::Stack,
        line,
    );
}

fn set_local(chunk: &mut Chunk, obj: u16, field: &str, value: u16, line: u32) {
    class_slots::emit_class_set(
        chunk,
        class_slots::ObjSource::Local(obj),
        &slot(field),
        class_slots::ValueSource::Local(value),
        line,
    );
}

fn set_const_str(chunk: &mut Chunk, obj: u16, field: &str, value: &str, line: u32) {
    class_slots::emit_class_set(
        chunk,
        class_slots::ObjSource::Local(obj),
        &slot(field),
        class_slots::ValueSource::ConstStr(value.to_string()),
        line,
    );
}

fn slot(field: &str) -> class_slots::ResolvedSlot {
    class_slots::resolve(
        &class_slots::ClassSlot::internal(field),
        &class_slots::PlainNames,
    )
}
