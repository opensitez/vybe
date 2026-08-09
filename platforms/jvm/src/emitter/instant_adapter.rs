//! Small Java `java.time` adapter backed by ECMA date helpers.

use vybe_ast::datetime::{DateTimePolicy, MonthIndexing, WeekdayBase};
use vybe_compiler::primitives::{
    collections, datetime,
    instructions::{core_wasm, host},
    ops,
};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn prop_get(chunks: &mut [Chunk], current: usize, obj: u16, key: &str, line: u32) {
    get(&mut chunks[current], obj, line);
    chunks[current].emit_string_const(key, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
}

fn prop_set(chunks: &mut [Chunk], current: usize, obj: u16, key: &str, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], obj, line);
    chunks[current].emit_string_const(key, line);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn make_instant_from_slots(chunks: &mut [Chunk], current: usize, sec: u16, nano: u16, line: u32) {
    host::emit(&mut chunks[current], "ecma:object", "new", 0, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    get(&mut chunks[current], sec, line);
    prop_set(chunks, current, out, "sec", line);
    get(&mut chunks[current], nano, line);
    prop_set(chunks, current, out, "nano", line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    prop_set(chunks, current, out, "offset", line);
    get(&mut chunks[current], out, line);
}

fn make_time_from_ms_slots(
    chunks: &mut [Chunk],
    current: usize,
    ms: u16,
    nano: u16,
    offset: u16,
    kind: &str,
    line: u32,
) {
    get(&mut chunks[current], ms, line);
    core_wasm::f64_const(&mut chunks[current], line, 1000.0);
    chunks[current].emit_op(Op::F64_DIV, line);
    host::emit(&mut chunks[current], "ecma:math", "floor", 1, line);
    let sec = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], sec, line);
    make_instant_from_slots(chunks, current, sec, nano, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    get(&mut chunks[current], offset, line);
    prop_set(chunks, current, out, "offset", line);
    chunks[current].emit_string_const(kind, line);
    prop_set(chunks, current, out, "kind", line);
    get(&mut chunks[current], out, line);
}

fn make_time_from_ms_kind_slot(
    chunks: &mut [Chunk],
    current: usize,
    ms: u16,
    nano: u16,
    offset: u16,
    kind: u16,
    line: u32,
) {
    get(&mut chunks[current], ms, line);
    core_wasm::f64_const(&mut chunks[current], line, 1000.0);
    chunks[current].emit_op(Op::F64_DIV, line);
    host::emit(&mut chunks[current], "ecma:math", "floor", 1, line);
    let sec = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], sec, line);
    make_instant_from_slots(chunks, current, sec, nano, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    get(&mut chunks[current], offset, line);
    prop_set(chunks, current, out, "offset", line);
    get(&mut chunks[current], kind, line);
    prop_set(chunks, current, out, "kind", line);
    get(&mut chunks[current], out, line);
}

fn zero_slot(chunks: &mut [Chunk], current: usize, line: u32) -> u16 {
    core_wasm::i32_const(&mut chunks[current], line, 0);
    let slot = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], slot, line);
    slot
}

fn date_utc_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    year: u16,
    month_one_based: u16,
    day: u16,
    hour: u16,
    minute: u16,
    second: u16,
    line: u32,
) {
    get(&mut chunks[current], year, line);
    get(&mut chunks[current], month_one_based, line);
    core_wasm::i32_const(&mut chunks[current], line, -1);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    get(&mut chunks[current], day, line);
    get(&mut chunks[current], hour, line);
    get(&mut chunks[current], minute, line);
    get(&mut chunks[current], second, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    host::emit(&mut chunks[current], "ecma:date", "UTC", 7, line);
}

fn emit_time_from_components(
    chunks: &mut [Chunk],
    current: usize,
    year: u16,
    month: u16,
    day: u16,
    hour: u16,
    minute: u16,
    second: u16,
    nano: u16,
    offset: u16,
    kind: &str,
    line: u32,
) {
    date_utc_from_slots(
        chunks, current, year, month, day, hour, minute, second, line,
    );
    emit_offset_seconds_from_value(chunks, current, offset, line);
    core_wasm::f64_const(&mut chunks[current], line, 1000.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    let ms = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], ms, line);
    make_time_from_ms_slots(chunks, current, ms, nano, offset, kind, line);
}

pub fn emit_local_date_of(chunks: &mut [Chunk], current: usize, line: u32) {
    let day = chunks[current].alloc_scratch(1);
    let month = chunks[current].alloc_scratch(1);
    let year = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], day, line);
    set(&mut chunks[current], month, line);
    set(&mut chunks[current], year, line);
    let zero = zero_slot(chunks, current, line);
    emit_time_from_components(
        chunks,
        current,
        year,
        month,
        day,
        zero,
        zero,
        zero,
        zero,
        zero,
        "LocalDate",
        line,
    );
}

pub fn emit_local_time_of(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let nano_arg = if argc >= 4 {
        let slot = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], slot, line);
        Some(slot)
    } else {
        None
    };
    let second_arg = if argc >= 3 {
        let slot = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], slot, line);
        Some(slot)
    } else {
        None
    };
    let minute = chunks[current].alloc_scratch(1);
    let hour = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], minute, line);
    set(&mut chunks[current], hour, line);
    let zero = zero_slot(chunks, current, line);
    let nano = nano_arg.unwrap_or(zero);
    let second = second_arg.unwrap_or(zero);
    core_wasm::i32_const(&mut chunks[current], line, 1970);
    let year = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], year, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    let month = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], month, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    let day = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], day, line);
    emit_time_from_components(
        chunks,
        current,
        year,
        month,
        day,
        hour,
        minute,
        second,
        nano,
        zero,
        "LocalTime",
        line,
    );
}

pub fn emit_local_datetime_of(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let nano_arg = if argc >= 7 {
        let slot = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], slot, line);
        Some(slot)
    } else {
        None
    };
    let second_arg = if argc >= 6 {
        let slot = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], slot, line);
        Some(slot)
    } else {
        None
    };
    let minute = chunks[current].alloc_scratch(1);
    let hour = chunks[current].alloc_scratch(1);
    let day = chunks[current].alloc_scratch(1);
    let month = chunks[current].alloc_scratch(1);
    let year = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], minute, line);
    set(&mut chunks[current], hour, line);
    set(&mut chunks[current], day, line);
    set(&mut chunks[current], month, line);
    set(&mut chunks[current], year, line);
    let zero = zero_slot(chunks, current, line);
    let nano = nano_arg.unwrap_or(zero);
    let second = second_arg.unwrap_or(zero);
    emit_time_from_components(
        chunks,
        current,
        year,
        month,
        day,
        hour,
        minute,
        second,
        nano,
        zero,
        "LocalDateTime",
        line,
    );
}

pub fn emit_local_date_parse(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_string_const("T00:00:00Z", line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    host::emit(&mut chunks[current], "ecma:date", "parse", 1, line);
    let ms = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], ms, line);
    let zero = zero_slot(chunks, current, line);
    make_time_from_ms_slots(chunks, current, ms, zero, zero, "LocalDate", line);
}

/// `YearMonth.parse("2024-05")` — a first-of-month LocalDate carries the
/// year+month surface the type exposes.
pub fn emit_year_month_parse(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_string_const("-01", line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    emit_local_date_parse(chunks, current, line);
}

/// `MonthDay.parse("--12-25")` — year 2000 (a leap year, so `--02-29`
/// parses), month/day from the `--MM-dd` tail.
pub fn emit_month_day_parse(chunks: &mut [Chunk], current: usize, line: u32) {
    let s = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], s, line);
    chunks[current].emit_string_const("2000-", line);
    get(&mut chunks[current], s, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    get(&mut chunks[current], s, line);
    vybe_compiler::primitives::strings::emit_length(&mut chunks[current], line);
    host::emit(&mut chunks[current], "ecma:string", "substring", 3, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    emit_local_date_parse(chunks, current, line);
}

pub fn emit_local_time_parse(chunks: &mut [Chunk], current: usize, line: u32) {
    let time = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], time, line);
    chunks[current].emit_string_const("1970-01-01T", line);
    // "08:00" (minute precision) is a valid LocalTime — the host parser
    // wants seconds, same normalization the datetime parse does.
    get(&mut chunks[current], time, line);
    vybe_compiler::primitives::strings::emit_length(&mut chunks[current], line);
    core_wasm::i32_const(&mut chunks[current], line, 5);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], time, line);
    chunks[current].emit_string_const(":00", line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], time, line);
    chunks[current].emit_end(line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_string_const("Z", line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    host::emit(&mut chunks[current], "ecma:date", "parse", 1, line);
    let ms = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], ms, line);
    let zero = zero_slot(chunks, current, line);
    make_time_from_ms_slots(chunks, current, ms, zero, zero, "LocalTime", line);
}

pub fn emit_local_datetime_parse(chunks: &mut [Chunk], current: usize, line: u32) {
    // `LocalDateTime.parse` accepts MINUTE precision ("2024-01-01T00:00");
    // the host date parser does not — every such value silently became NaN
    // (and compareTo answered 1 for two equal parses). Normalize to seconds.
    let s = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], s, line);
    get(&mut chunks[current], s, line);
    vybe_compiler::primitives::strings::emit_length(&mut chunks[current], line);
    core_wasm::i32_const(&mut chunks[current], line, 16);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], s, line);
    chunks[current].emit_string_const(":00", line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_string_const("Z", line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    host::emit(&mut chunks[current], "ecma:date", "parse", 1, line);
    let ms = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], ms, line);
    let zero = zero_slot(chunks, current, line);
    make_time_from_ms_slots(chunks, current, ms, zero, zero, "LocalDateTime", line);
}

pub fn emit_offset_datetime_of(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 2 {
        let offset = chunks[current].alloc_scratch(1);
        let local = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], offset, line);
        set(&mut chunks[current], local, line);
        emit_epoch_milli_from_slot(chunks, current, local, line);
        emit_offset_seconds_from_value(chunks, current, offset, line);
        core_wasm::f64_const(&mut chunks[current], line, 1000.0);
        chunks[current].emit_op(Op::F64_MUL, line);
        chunks[current].emit_op(Op::F64_SUB, line);
        let ms = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], ms, line);
        prop_get(chunks, current, local, "nano", line);
        let nano = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], nano, line);
        make_time_from_ms_slots(chunks, current, ms, nano, offset, "OffsetDateTime", line);
        return;
    }

    let offset = chunks[current].alloc_scratch(1);
    let nano = chunks[current].alloc_scratch(1);
    let second = chunks[current].alloc_scratch(1);
    let minute = chunks[current].alloc_scratch(1);
    let hour = chunks[current].alloc_scratch(1);
    let day = chunks[current].alloc_scratch(1);
    let month = chunks[current].alloc_scratch(1);
    let year = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], offset, line);
    set(&mut chunks[current], nano, line);
    set(&mut chunks[current], second, line);
    set(&mut chunks[current], minute, line);
    set(&mut chunks[current], hour, line);
    set(&mut chunks[current], day, line);
    set(&mut chunks[current], month, line);
    set(&mut chunks[current], year, line);
    emit_time_from_components(
        chunks,
        current,
        year,
        month,
        day,
        hour,
        minute,
        second,
        nano,
        offset,
        "OffsetDateTime",
        line,
    );
}

pub fn emit_offset_datetime_of_instant(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_with_offset(chunks, current, line);
}

pub fn emit_offset_datetime_parse(chunks: &mut [Chunk], current: usize, line: u32) {
    let text = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], text, line);
    get(&mut chunks[current], text, line);
    host::emit(&mut chunks[current], "ecma:date", "parse", 1, line);
    let ms = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], ms, line);

    get(&mut chunks[current], text, line);
    core_wasm::i32_const(&mut chunks[current], line, 19);
    core_wasm::i32_const(&mut chunks[current], line, 20);
    host::emit(&mut chunks[current], "ecma:string", "substring", 3, line);
    chunks[current].emit_string_const("Z", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("Z", line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], text, line);
    core_wasm::i32_const(&mut chunks[current], line, 19);
    core_wasm::i32_const(&mut chunks[current], line, 25);
    host::emit(&mut chunks[current], "ecma:string", "substring", 3, line);
    chunks[current].emit_end(line);
    let offset = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], offset, line);
    let zero = zero_slot(chunks, current, line);
    make_time_from_ms_slots(chunks, current, ms, zero, offset, "OffsetDateTime", line);
}

pub fn emit_time_with_offset_same_local(chunks: &mut [Chunk], current: usize, line: u32) {
    let offset = chunks[current].alloc_scratch(1);
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], offset, line);
    set(&mut chunks[current], inst, line);
    emit_epoch_milli_from_slot(chunks, current, inst, line);
    prop_get(chunks, current, inst, "offset", line);
    let old_offset = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], old_offset, line);
    emit_offset_seconds_from_value(chunks, current, old_offset, line);
    core_wasm::f64_const(&mut chunks[current], line, 1000.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    emit_offset_seconds_from_value(chunks, current, offset, line);
    core_wasm::f64_const(&mut chunks[current], line, 1000.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    let ms = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], ms, line);
    prop_get(chunks, current, inst, "nano", line);
    let nano = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], nano, line);
    make_time_from_ms_slots(chunks, current, ms, nano, offset, "OffsetDateTime", line);
}

fn set_zone_from_slot(chunks: &mut [Chunk], current: usize, out: u16, zone: u16, line: u32) {
    get(&mut chunks[current], zone, line);
    prop_set(chunks, current, out, "zone", line);
}

pub fn emit_zoned_datetime_of(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 2 {
        emit_offset_datetime_of(chunks, current, argc, line);
        let out = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], out, line);
        prop_get(chunks, current, out, "offset", line);
        let zone = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], zone, line);
        set_zone_from_slot(chunks, current, out, zone, line);
        get(&mut chunks[current], out, line);
        return;
    }
    let zone = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], zone, line);
    get(&mut chunks[current], zone, line);
    // Reuse OffsetDateTime.of with the zone as its offset model.
    emit_offset_datetime_of(chunks, current, argc, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    set_zone_from_slot(chunks, current, out, zone, line);
    get(&mut chunks[current], out, line);
}

pub fn emit_zoned_datetime_of_instant(chunks: &mut [Chunk], current: usize, line: u32) {
    let zone = chunks[current].alloc_scratch(1);
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], zone, line);
    set(&mut chunks[current], inst, line);
    get(&mut chunks[current], inst, line);
    get(&mut chunks[current], zone, line);
    emit_with_offset(chunks, current, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    set_zone_from_slot(chunks, current, out, zone, line);
    get(&mut chunks[current], out, line);
}

pub fn emit_zoned_datetime_of_strict(chunks: &mut [Chunk], current: usize, line: u32) {
    let zone = chunks[current].alloc_scratch(1);
    let offset = chunks[current].alloc_scratch(1);
    let local = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], zone, line);
    set(&mut chunks[current], offset, line);
    set(&mut chunks[current], local, line);
    get(&mut chunks[current], local, line);
    get(&mut chunks[current], offset, line);
    emit_offset_datetime_of(chunks, current, 2, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    set_zone_from_slot(chunks, current, out, zone, line);
    get(&mut chunks[current], out, line);
}

pub fn emit_zoned_datetime_parse(chunks: &mut [Chunk], current: usize, line: u32) {
    let text = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], text, line);
    get(&mut chunks[current], text, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 25);
    host::emit(&mut chunks[current], "ecma:string", "substring", 3, line);
    let parse_text = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], parse_text, line);
    get(&mut chunks[current], text, line);
    core_wasm::i32_const(&mut chunks[current], line, 19);
    core_wasm::i32_const(&mut chunks[current], line, 20);
    host::emit(&mut chunks[current], "ecma:string", "substring", 3, line);
    chunks[current].emit_string_const("Z", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], text, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], parse_text, line);
    chunks[current].emit_end(line);
    emit_offset_datetime_parse(chunks, current, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    // A `[Europe/Paris]` suffix IS the zone — the offset only stands in
    // when the bracket form is absent (`value.zone.id` printed the offset).
    let bracket = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], text, line);
    chunks[current].emit_string_const("[", line);
    vybe_compiler::primitives::strings::emit_index_of(&mut chunks[current], line);
    set(&mut chunks[current], bracket, line);
    get(&mut chunks[current], bracket, line);
    core_wasm::f64_const(&mut chunks[current], line, 0.0);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if_value(line);
    prop_get(chunks, current, out, "offset", line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], text, line);
    get(&mut chunks[current], bracket, line);
    core_wasm::f64_const(&mut chunks[current], line, 1.0);
    chunks[current].emit_op(Op::F64_ADD, line);
    get(&mut chunks[current], text, line);
    vybe_compiler::primitives::strings::emit_length(&mut chunks[current], line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    host::emit(&mut chunks[current], "ecma:string", "substring", 3, line);
    chunks[current].emit_end(line);
    let zone = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], zone, line);
    set_zone_from_slot(chunks, current, out, zone, line);
    get(&mut chunks[current], out, line);
}

pub fn emit_get_zone(chunks: &mut [Chunk], current: usize, line: u32) {
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], inst, line);
    prop_get(chunks, current, inst, "zone", line);
    let zone = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], zone, line);
    get(&mut chunks[current], zone, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("undefined", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    prop_get(chunks, current, inst, "offset", line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], zone, line);
    chunks[current].emit_end(line);
}

pub fn emit_with_zone_same_instant(chunks: &mut [Chunk], current: usize, line: u32) {
    let zone = chunks[current].alloc_scratch(1);
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], zone, line);
    set(&mut chunks[current], inst, line);
    get(&mut chunks[current], inst, line);
    get(&mut chunks[current], zone, line);
    emit_with_offset(chunks, current, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    set_zone_from_slot(chunks, current, out, zone, line);
    get(&mut chunks[current], out, line);
}

pub fn emit_with_zone_same_local(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_time_with_offset_same_local(chunks, current, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    prop_get(chunks, current, out, "offset", line);
    let zone = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], zone, line);
    set_zone_from_slot(chunks, current, out, zone, line);
    get(&mut chunks[current], out, line);
}

pub fn emit_overlap_offset(chunks: &mut [Chunk], current: usize, hours: i32, line: u32) {
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], inst, line);
    get(&mut chunks[current], inst, line);
    core_wasm::i32_const(&mut chunks[current], line, hours);
    emit_zone_offset_hours(chunks, current, line);
    emit_with_offset(chunks, current, line);
}

pub fn emit_of_epoch_second(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let nano = chunks[current].alloc_scratch(1);
    if argc == 1 {
        core_wasm::i32_const(&mut chunks[current], line, 0);
    }
    set(&mut chunks[current], nano, line);
    let sec = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], sec, line);
    make_instant_from_slots(chunks, current, sec, nano, line);
}

pub fn emit_of_epoch_milli(chunks: &mut [Chunk], current: usize, line: u32) {
    let ms = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], ms, line);
    get(&mut chunks[current], ms, line);
    core_wasm::f64_const(&mut chunks[current], line, 1000.0);
    chunks[current].emit_op(Op::F64_DIV, line);
    host::emit(&mut chunks[current], "ecma:math", "floor", 1, line);
    let sec = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], sec, line);
    get(&mut chunks[current], ms, line);
    get(&mut chunks[current], sec, line);
    core_wasm::f64_const(&mut chunks[current], line, 1000.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    core_wasm::f64_const(&mut chunks[current], line, 1_000_000.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    let nano = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], nano, line);
    make_instant_from_slots(chunks, current, sec, nano, line);
}

pub fn emit_parse(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:date", "parse", 1, line);
    emit_of_epoch_milli(chunks, current, line);
}

pub fn emit_get_epoch_second(chunks: &mut [Chunk], current: usize, line: u32) {
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], inst, line);
    prop_get(chunks, current, inst, "sec", line);
}

pub fn emit_get_nano(chunks: &mut [Chunk], current: usize, line: u32) {
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], inst, line);
    prop_get(chunks, current, inst, "nano", line);
}

fn emit_epoch_milli_from_slot(chunks: &mut [Chunk], current: usize, inst: u16, line: u32) {
    prop_get(chunks, current, inst, "sec", line);
    core_wasm::f64_const(&mut chunks[current], line, 1000.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    prop_get(chunks, current, inst, "nano", line);
    core_wasm::f64_const(&mut chunks[current], line, 1_000_000.0);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_ADD, line);
}

pub fn emit_to_epoch_milli(chunks: &mut [Chunk], current: usize, line: u32) {
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], inst, line);
    emit_epoch_milli_from_slot(chunks, current, inst, line);
}

pub fn emit_plus_seconds(chunks: &mut [Chunk], current: usize, sign: f64, line: u32) {
    let delta = chunks[current].alloc_scratch(1);
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], delta, line);
    set(&mut chunks[current], inst, line);
    prop_get(chunks, current, inst, "sec", line);
    get(&mut chunks[current], delta, line);
    core_wasm::f64_const(&mut chunks[current], line, sign);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    let sec = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], sec, line);
    prop_get(chunks, current, inst, "nano", line);
    let nano = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], nano, line);
    make_instant_from_slots(chunks, current, sec, nano, line);
}

pub fn emit_plus_millis(chunks: &mut [Chunk], current: usize, sign: f64, line: u32) {
    let delta = chunks[current].alloc_scratch(1);
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], delta, line);
    set(&mut chunks[current], inst, line);
    emit_epoch_milli_from_slot(chunks, current, inst, line);
    get(&mut chunks[current], delta, line);
    core_wasm::f64_const(&mut chunks[current], line, sign);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    emit_of_epoch_milli(chunks, current, line);
}

pub fn emit_plus_nanos(chunks: &mut [Chunk], current: usize, sign: f64, line: u32) {
    let delta = chunks[current].alloc_scratch(1);
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], delta, line);
    set(&mut chunks[current], inst, line);
    prop_get(chunks, current, inst, "sec", line);
    let sec = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], sec, line);
    prop_get(chunks, current, inst, "nano", line);
    get(&mut chunks[current], delta, line);
    core_wasm::f64_const(&mut chunks[current], line, sign);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    let nano = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], nano, line);
    make_instant_from_slots(chunks, current, sec, nano, line);
}

pub fn emit_compare(chunks: &mut [Chunk], current: usize, line: u32) {
    let other = chunks[current].alloc_scratch(1);
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], other, line);
    set(&mut chunks[current], inst, line);
    emit_epoch_milli_from_slot(chunks, current, inst, line);
    emit_epoch_milli_from_slot(chunks, current, other, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_else(line);
    emit_epoch_milli_from_slot(chunks, current, inst, line);
    emit_epoch_milli_from_slot(chunks, current, other, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    core_wasm::i32_const(&mut chunks[current], line, -1);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_is_before_after(chunks: &mut [Chunk], current: usize, after: bool, line: u32) {
    emit_compare(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    if after {
        vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    } else {
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    }
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_compare(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], inst, line);
    emit_epoch_milli_from_slot(chunks, current, inst, line);
    host::emit(&mut chunks[current], "ecma:date", "new", 1, line);
    host::emit(&mut chunks[current], "ecma:date", "toISOString", 1, line);
    chunks[current].emit_string_const(".000Z", line);
    chunks[current].emit_string_const("Z", line);
    host::emit(&mut chunks[current], "ecma:string", "replace", 3, line);
}

pub fn emit_duration_minutes(chunks: &mut [Chunk], current: usize, line: u32) {
    core_wasm::f64_const(&mut chunks[current], line, 60.0);
    chunks[current].emit_op(Op::F64_MUL, line);
}

pub fn emit_duration_hours(chunks: &mut [Chunk], current: usize, line: u32) {
    core_wasm::f64_const(&mut chunks[current], line, 3600.0);
    chunks[current].emit_op(Op::F64_MUL, line);
}

pub fn emit_duration_seconds(_chunks: &mut [Chunk], _current: usize, _line: u32) {}

pub fn emit_duration_parse(chunks: &mut [Chunk], current: usize, line: u32) {
    let text = chunks[current].alloc_scratch(1);
    let body = chunks[current].alloc_scratch(1);
    let parts = chunks[current].alloc_scratch(1);
    let hours = chunks[current].alloc_scratch(1);
    let minutes = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], text, line);
    get(&mut chunks[current], text, line);
    chunks[current].emit_string_const("PT", line);
    chunks[current].emit_string_const("", line);
    host::emit(&mut chunks[current], "ecma:string", "replace", 3, line);
    set(&mut chunks[current], body, line);
    get(&mut chunks[current], body, line);
    chunks[current].emit_string_const("H", line);
    host::emit(&mut chunks[current], "ecma:string", "split", 2, line);
    set(&mut chunks[current], parts, line);
    get(&mut chunks[current], parts, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 10);
    host::emit(&mut chunks[current], "ecma:number", "parseInt", 2, line);
    set(&mut chunks[current], hours, line);
    get(&mut chunks[current], parts, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_string_const("M", line);
    chunks[current].emit_string_const("", line);
    host::emit(&mut chunks[current], "ecma:string", "replace", 3, line);
    core_wasm::i32_const(&mut chunks[current], line, 10);
    host::emit(&mut chunks[current], "ecma:number", "parseInt", 2, line);
    set(&mut chunks[current], minutes, line);
    get(&mut chunks[current], hours, line);
    core_wasm::f64_const(&mut chunks[current], line, 3600.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    get(&mut chunks[current], minutes, line);
    core_wasm::f64_const(&mut chunks[current], line, 60.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
}

pub fn emit_duration_to_hours(chunks: &mut [Chunk], current: usize, line: u32) {
    core_wasm::f64_const(&mut chunks[current], line, 3600.0);
    chunks[current].emit_op(Op::F64_DIV, line);
    host::emit(&mut chunks[current], "ecma:math", "floor", 1, line);
}

pub fn emit_duration_to_minutes(chunks: &mut [Chunk], current: usize, line: u32) {
    core_wasm::f64_const(&mut chunks[current], line, 60.0);
    chunks[current].emit_op(Op::F64_DIV, line);
}

pub fn emit_duration_to_millis(chunks: &mut [Chunk], current: usize, line: u32) {
    core_wasm::f64_const(&mut chunks[current], line, 1000.0);
    chunks[current].emit_op(Op::F64_MUL, line);
}

pub fn emit_duration_plus_hours(chunks: &mut [Chunk], current: usize, sign: f64, line: u32) {
    let amount = chunks[current].alloc_scratch(1);
    let duration = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], amount, line);
    set(&mut chunks[current], duration, line);
    get(&mut chunks[current], duration, line);
    get(&mut chunks[current], amount, line);
    core_wasm::f64_const(&mut chunks[current], line, 3600.0 * sign);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
}

pub fn emit_duration_multiplied_by(chunks: &mut [Chunk], current: usize, line: u32) {
    let factor = chunks[current].alloc_scratch(1);
    let duration = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], factor, line);
    set(&mut chunks[current], duration, line);
    get(&mut chunks[current], duration, line);
    get(&mut chunks[current], factor, line);
    chunks[current].emit_op(Op::F64_MUL, line);
}

pub fn emit_duration_negated(chunks: &mut [Chunk], current: usize, line: u32) {
    core_wasm::f64_const(&mut chunks[current], line, -1.0);
    chunks[current].emit_op(Op::F64_MUL, line);
}

pub fn emit_duration_is_zero(chunks: &mut [Chunk], current: usize, line: u32) {
    core_wasm::f64_const(&mut chunks[current], line, 0.0);
    chunks[current].emit_op(Op::F64_EQ, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_duration_plus_minutes(chunks: &mut [Chunk], current: usize, sign: f64, line: u32) {
    let amount = chunks[current].alloc_scratch(1);
    let duration = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], amount, line);
    set(&mut chunks[current], duration, line);
    get(&mut chunks[current], duration, line);
    get(&mut chunks[current], amount, line);
    core_wasm::f64_const(&mut chunks[current], line, 60.0 * sign);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
}

pub fn emit_duration_between(chunks: &mut [Chunk], current: usize, line: u32) {
    let end = chunks[current].alloc_scratch(1);
    let start = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], end, line);
    set(&mut chunks[current], start, line);
    prop_get(chunks, current, end, "sec", line);
    prop_get(chunks, current, start, "sec", line);
    chunks[current].emit_op(Op::F64_SUB, line);
}

pub fn emit_chrono_between(chunks: &mut [Chunk], current: usize, seconds: f64, line: u32) {
    emit_duration_between(chunks, current, line);
    core_wasm::f64_const(&mut chunks[current], line, seconds);
    chunks[current].emit_op(Op::F64_DIV, line);
    host::emit(&mut chunks[current], "ecma:math", "floor", 1, line);
}

pub fn emit_clock_fixed(chunks: &mut [Chunk], current: usize, line: u32) {
    let zone = chunks[current].alloc_scratch(1);
    let instant = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], zone, line);
    set(&mut chunks[current], instant, line);
    host::emit(&mut chunks[current], "ecma:object", "new", 0, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    get(&mut chunks[current], instant, line);
    prop_set(chunks, current, out, "instant", line);
    get(&mut chunks[current], zone, line);
    prop_set(chunks, current, out, "zone", line);
    get(&mut chunks[current], out, line);
}

pub fn emit_instant_now(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 1 {
        let clock = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], clock, line);
        prop_get(chunks, current, clock, "instant", line);
        return;
    }
    host::emit(&mut chunks[current], "ecma:date", "now", 0, line);
    let ms = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], ms, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    let nano = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], nano, line);
    get(&mut chunks[current], ms, line);
    core_wasm::f64_const(&mut chunks[current], line, 1000.0);
    chunks[current].emit_op(Op::F64_DIV, line);
    host::emit(&mut chunks[current], "ecma:math", "floor", 1, line);
    let sec = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], sec, line);
    make_instant_from_slots(chunks, current, sec, nano, line);
}

fn make_period_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    years: u16,
    months: u16,
    days: u16,
    line: u32,
) {
    host::emit(&mut chunks[current], "ecma:object", "new", 0, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    get(&mut chunks[current], years, line);
    prop_set(chunks, current, out, "years", line);
    get(&mut chunks[current], months, line);
    prop_set(chunks, current, out, "months", line);
    get(&mut chunks[current], days, line);
    prop_set(chunks, current, out, "days", line);
    chunks[current].emit_string_const("Period", line);
    prop_set(chunks, current, out, "kind", line);
    get(&mut chunks[current], out, line);
}

pub fn emit_period_of_days(chunks: &mut [Chunk], current: usize, line: u32) {
    let days = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], days, line);
    let zero = zero_slot(chunks, current, line);
    make_period_from_slots(chunks, current, zero, zero, days, line);
}

pub fn emit_period_of_months(chunks: &mut [Chunk], current: usize, line: u32) {
    let months = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], months, line);
    let zero = zero_slot(chunks, current, line);
    make_period_from_slots(chunks, current, zero, months, zero, line);
}

pub fn emit_period_get_days(chunks: &mut [Chunk], current: usize, line: u32) {
    let period = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], period, line);
    prop_get(chunks, current, period, "days", line);
}

pub fn emit_period_get_months(chunks: &mut [Chunk], current: usize, line: u32) {
    let period = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], period, line);
    prop_get(chunks, current, period, "months", line);
}

pub fn emit_period_get_years(chunks: &mut [Chunk], current: usize, line: u32) {
    let period = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], period, line);
    prop_get(chunks, current, period, "years", line);
}

pub fn emit_period_between(chunks: &mut [Chunk], current: usize, line: u32) {
    // Java's CALENDAR decomposition, not a day count: 2024-01-01 →
    // 2024-03-11 is 2 months 10 days (the flat `/86400` answered 0y 0m 70d).
    let end = chunks[current].alloc_scratch(1);
    let start = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], end, line);
    set(&mut chunks[current], start, line);

    fn component(chunks: &mut [Chunk], current: usize, slot: u16, getter: &'static str, line: u32) {
        emit_epoch_milli_from_slot(chunks, current, slot, line);
        host::emit(&mut chunks[current], "ecma:date", "new", 1, line);
        host::emit(&mut chunks[current], "ecma:date", getter, 1, line);
    }

    let months = chunks[current].alloc_scratch(1);
    let days = chunks[current].alloc_scratch(1);
    let years = chunks[current].alloc_scratch(1);

    // months = (y2 - y1) * 12 + (m2 - m1)
    component(chunks, current, end, "getUTCFullYear", line);
    component(chunks, current, start, "getUTCFullYear", line);
    chunks[current].emit_op(Op::F64_SUB, line);
    core_wasm::f64_const(&mut chunks[current], line, 12.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    component(chunks, current, end, "getUTCMonth", line);
    component(chunks, current, start, "getUTCMonth", line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    set(&mut chunks[current], months, line);

    // days = d2 - d1
    component(chunks, current, end, "getUTCDate", line);
    component(chunks, current, start, "getUTCDate", line);
    chunks[current].emit_op(Op::F64_SUB, line);
    set(&mut chunks[current], days, line);

    // Borrow a month when days went negative: the borrowed month's length
    // is the month BEFORE the end month (`Date.UTC(y2, m2, 0)` is its last
    // day).
    get(&mut chunks[current], days, line);
    core_wasm::f64_const(&mut chunks[current], line, 0.0);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], months, line);
    core_wasm::f64_const(&mut chunks[current], line, 1.0);
    chunks[current].emit_op(Op::F64_SUB, line);
    set(&mut chunks[current], months, line);
    get(&mut chunks[current], days, line);
    component(chunks, current, end, "getUTCFullYear", line);
    component(chunks, current, end, "getUTCMonth", line);
    for _ in 0..5 {
        core_wasm::f64_const(&mut chunks[current], line, 0.0);
    }
    host::emit(&mut chunks[current], "ecma:date", "UTC", 7, line);
    host::emit(&mut chunks[current], "ecma:date", "new", 1, line);
    host::emit(&mut chunks[current], "ecma:date", "getUTCDate", 1, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    set(&mut chunks[current], days, line);
    chunks[current].emit_end(line);

    // years = trunc(months / 12); months -= years * 12
    get(&mut chunks[current], months, line);
    core_wasm::f64_const(&mut chunks[current], line, 12.0);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_TRUNC, line);
    set(&mut chunks[current], years, line);
    get(&mut chunks[current], months, line);
    get(&mut chunks[current], years, line);
    core_wasm::f64_const(&mut chunks[current], line, 12.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    set(&mut chunks[current], months, line);

    make_period_from_slots(chunks, current, years, months, days, line);
}

pub fn emit_with_offset(chunks: &mut [Chunk], current: usize, line: u32) {
    let offset = chunks[current].alloc_scratch(1);
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], offset, line);
    set(&mut chunks[current], inst, line);
    prop_get(chunks, current, inst, "sec", line);
    let sec = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], sec, line);
    prop_get(chunks, current, inst, "nano", line);
    let nano = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], nano, line);
    make_instant_from_slots(chunks, current, sec, nano, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    get(&mut chunks[current], offset, line);
    prop_set(chunks, current, out, "offset", line);
    get(&mut chunks[current], out, line);
}

pub fn emit_zone_offset_hours(chunks: &mut [Chunk], current: usize, line: u32) {
    let hours = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], hours, line);
    get(&mut chunks[current], hours, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("-", line);
    get(&mut chunks[current], hours, line);
    core_wasm::f64_const(&mut chunks[current], line, -1.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("+", line);
    get(&mut chunks[current], hours, line);
    chunks[current].emit_end(line);
    let abs_hours = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], abs_hours, line);
    let sign = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], sign, line);
    get(&mut chunks[current], abs_hours, line);
    core_wasm::i32_const(&mut chunks[current], line, 10);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], sign, line);
    chunks[current].emit_string_const("0", line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], sign, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], abs_hours, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    host::emit(&mut chunks[current], "ecma:number", "toFixed", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_string_const(":00", line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
}

pub fn emit_zone_id_utc(_chunks: &mut [Chunk], _current: usize, _line: u32) {
    // ZoneId.of(id) keeps the textual id.
}

pub fn emit_zone_id_system_default(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_string_const("UTC", line);
}

pub fn emit_zone_id_short_ids(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:map", "new", 0, line);
    let map = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], map, line);
    for (key, value) in [
        ("UTC", "UTC"),
        ("EST", "America/New_York"),
        ("PST", "America/Los_Angeles"),
        ("CST", "America/Chicago"),
    ] {
        get(&mut chunks[current], map, line);
        chunks[current].emit_string_const(key, line);
        chunks[current].emit_string_const(value, line);
        host::emit(&mut chunks[current], "ecma:map", "set", 3, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    get(&mut chunks[current], map, line);
}

pub fn emit_zone_id_from(_chunks: &mut [Chunk], _current: usize, _line: u32) {
    // ZoneId.from(ZoneOffset) is identical for the string representation.
}

pub fn emit_zone_id_of_offset(chunks: &mut [Chunk], current: usize, line: u32) {
    let offset = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], offset, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], offset, line);
}

pub fn emit_zone_normalized(chunks: &mut [Chunk], current: usize, line: u32) {
    let zone = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], zone, line);
    get(&mut chunks[current], zone, line);
    chunks[current].emit_string_const("UTC", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("Z", line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], zone, line);
    chunks[current].emit_end(line);
}

pub fn emit_zone_display_name(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc.saturating_sub(1) {
        chunks[current].emit_op(Op::DROP, line);
    }
}

pub fn emit_zone_rules_fixed(chunks: &mut [Chunk], current: usize, line: u32) {
    let zone = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], zone, line);
    get(&mut chunks[current], zone, line);
    chunks[current].emit_string_const("Europe/Paris", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_get_offset(chunks: &mut [Chunk], current: usize, line: u32) {
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], inst, line);
    prop_get(chunks, current, inst, "offset", line);
}

fn emit_offset_seconds_from_value(
    chunks: &mut [Chunk],
    current: usize,
    value_slot: u16,
    line: u32,
) {
    get(&mut chunks[current], value_slot, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("string", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value_slot, line);
    chunks[current].emit_string_const("Z", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    get(&mut chunks[current], value_slot, line);
    chunks[current].emit_string_const("UTC", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if_value(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value_slot, line);
    chunks[current].emit_string_const("Asia/Tokyo", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    core_wasm::i32_const(&mut chunks[current], line, 32400);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value_slot, line);
    chunks[current].emit_string_const("Europe/Paris", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    core_wasm::i32_const(&mut chunks[current], line, 7200);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    host::emit(&mut chunks[current], "ecma:string", "substring", 3, line);
    chunks[current].emit_string_const("-", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    core_wasm::f64_const(&mut chunks[current], line, -1.0);
    chunks[current].emit_else(line);
    core_wasm::f64_const(&mut chunks[current], line, 1.0);
    chunks[current].emit_end(line);
    let sign = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], sign, line);
    get(&mut chunks[current], value_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    core_wasm::i32_const(&mut chunks[current], line, 3);
    host::emit(&mut chunks[current], "ecma:string", "substring", 3, line);
    core_wasm::i32_const(&mut chunks[current], line, 10);
    host::emit(&mut chunks[current], "ecma:number", "parseInt", 2, line);
    core_wasm::f64_const(&mut chunks[current], line, 3600.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    get(&mut chunks[current], value_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 4);
    core_wasm::i32_const(&mut chunks[current], line, 6);
    host::emit(&mut chunks[current], "ecma:string", "substring", 3, line);
    core_wasm::i32_const(&mut chunks[current], line, 10);
    host::emit(&mut chunks[current], "ecma:number", "parseInt", 2, line);
    core_wasm::f64_const(&mut chunks[current], line, 60.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    get(&mut chunks[current], sign, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value_slot, line);
    chunks[current].emit_end(line);
}

pub fn emit_get_total_seconds(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    emit_offset_seconds_from_value(chunks, current, value, line);
}

pub fn emit_zone_rules_get_offset(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    let zone = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], zone, line);
    get(&mut chunks[current], zone, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    host::emit(&mut chunks[current], "ecma:string", "substring", 3, line);
    chunks[current].emit_string_const("+", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    get(&mut chunks[current], zone, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    host::emit(&mut chunks[current], "ecma:string", "substring", 3, line);
    chunks[current].emit_string_const("-", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], zone, line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("Z", line);
    chunks[current].emit_end(line);
}

pub fn emit_zone_compare_to(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(
        &mut chunks[current],
        "ecma:string",
        "localeCompare",
        2,
        line,
    );
}

pub fn emit_zone_hash_code(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
}

pub fn emit_component(
    chunks: &mut [Chunk],
    current: usize,
    method: &'static str,
    add_one: bool,
    line: u32,
) {
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], inst, line);
    emit_epoch_milli_from_slot(chunks, current, inst, line);
    prop_get(chunks, current, inst, "offset", line);
    let offset = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], offset, line);
    emit_offset_seconds_from_value(chunks, current, offset, line);
    core_wasm::f64_const(&mut chunks[current], line, 1000.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    host::emit(&mut chunks[current], "ecma:date", "new", 1, line);
    host::emit(&mut chunks[current], "ecma:date", method, 1, line);
    if add_one {
        core_wasm::i32_const(&mut chunks[current], line, 1);
        vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    }
}

/// The full ISO instant string of the value's LOCAL wall time
/// (`2024-07-03T15:00:00.000Z`) — the slice each `toString` wants.
fn emit_iso_from_value(chunks: &mut [Chunk], current: usize, line: u32) {
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], inst, line);
    emit_epoch_milli_from_slot(chunks, current, inst, line);
    prop_get(chunks, current, inst, "offset", line);
    let offset = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], offset, line);
    emit_offset_seconds_from_value(chunks, current, offset, line);
    core_wasm::f64_const(&mut chunks[current], line, 1000.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    host::emit(&mut chunks[current], "ecma:date", "new", 1, line);
    host::emit(&mut chunks[current], "ecma:date", "toISOString", 1, line);
}

fn emit_iso_slice(chunks: &mut [Chunk], current: usize, from: i32, to: i32, line: u32) {
    core_wasm::i32_const(&mut chunks[current], line, from);
    core_wasm::i32_const(&mut chunks[current], line, to);
    host::emit(&mut chunks[current], "ecma:string", "substring", 3, line);
}

pub fn emit_local_date_string(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_iso_from_value(chunks, current, line);
    emit_iso_slice(chunks, current, 0, 10, line);
}

/// True when the ISO string in `iso` has `:00` seconds — Java's
/// `toString` omits the seconds field entirely then.
fn emit_iso_seconds_are_zero(chunks: &mut [Chunk], current: usize, iso: u16, line: u32) {
    get(&mut chunks[current], iso, line);
    emit_iso_slice(chunks, current, 17, 19, line);
    chunks[current].emit_string_const("00", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
}

/// `LocalDateTime.toString()` — `2024-07-03T15:00`, seconds only when nonzero.
pub fn emit_local_datetime_string(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_iso_from_value(chunks, current, line);
    let iso = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], iso, line);
    emit_iso_seconds_are_zero(chunks, current, iso, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], iso, line);
    emit_iso_slice(chunks, current, 0, 16, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], iso, line);
    emit_iso_slice(chunks, current, 0, 19, line);
    chunks[current].emit_end(line);
}

/// `LocalTime.toString()` — `02:05` / `09:30:45`; the ISO round-trip already
/// wraps past midnight.
pub fn emit_time_of_day_string(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_iso_from_value(chunks, current, line);
    let iso = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], iso, line);
    emit_iso_seconds_are_zero(chunks, current, iso, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], iso, line);
    emit_iso_slice(chunks, current, 11, 16, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], iso, line);
    emit_iso_slice(chunks, current, 11, 19, line);
    chunks[current].emit_end(line);
}

pub fn emit_time_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_local_date_string(chunks, current, line);
}

pub fn emit_time_format(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    emit_time_to_string(chunks, current, line);
}

pub fn emit_time_plus_unit(
    chunks: &mut [Chunk],
    current: usize,
    sign: f64,
    seconds_per_unit: f64,
    line: u32,
) {
    let delta = chunks[current].alloc_scratch(1);
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], delta, line);
    set(&mut chunks[current], inst, line);
    emit_epoch_milli_from_slot(chunks, current, inst, line);
    get(&mut chunks[current], delta, line);
    core_wasm::f64_const(&mut chunks[current], line, sign * seconds_per_unit * 1000.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    let ms = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], ms, line);
    prop_get(chunks, current, inst, "nano", line);
    let nano = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], nano, line);
    prop_get(chunks, current, inst, "offset", line);
    let offset = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], offset, line);
    make_time_from_ms_slots(chunks, current, ms, nano, offset, "LocalDateTime", line);
}

pub fn emit_time_plus_months(chunks: &mut [Chunk], current: usize, sign: f64, line: u32) {
    let delta = chunks[current].alloc_scratch(1);
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], delta, line);
    set(&mut chunks[current], inst, line);
    emit_epoch_milli_from_slot(chunks, current, inst, line);
    prop_get(chunks, current, inst, "offset", line);
    let offset = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], offset, line);
    get(&mut chunks[current], offset, line);
    core_wasm::f64_const(&mut chunks[current], line, 1000.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    get(&mut chunks[current], delta, line);
    core_wasm::f64_const(&mut chunks[current], line, sign);
    chunks[current].emit_op(Op::F64_MUL, line);
    datetime::emit_add_months(&mut chunks[current], DateTimePolicy::ISO, line);
    get(&mut chunks[current], offset, line);
    core_wasm::f64_const(&mut chunks[current], line, 1000.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    let ms = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], ms, line);
    prop_get(chunks, current, inst, "nano", line);
    let nano = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], nano, line);
    prop_get(chunks, current, inst, "kind", line);
    let kind = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], kind, line);
    make_time_from_ms_kind_slot(chunks, current, ms, nano, offset, kind, line);
}

pub fn emit_time_with_year_or_month(
    chunks: &mut [Chunk],
    current: usize,
    set_year: bool,
    line: u32,
) {
    let value = chunks[current].alloc_scratch(1);
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], inst, line);

    emit_epoch_milli_from_slot(chunks, current, inst, line);
    prop_get(chunks, current, inst, "offset", line);
    let offset = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], offset, line);
    get(&mut chunks[current], offset, line);
    core_wasm::f64_const(&mut chunks[current], line, 1000.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    let local_ms = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], local_ms, line);

    get(&mut chunks[current], local_ms, line);
    host::emit(&mut chunks[current], "ecma:date", "new", 1, line);
    let date = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], date, line);

    if set_year {
        get(&mut chunks[current], value, line);
        get(&mut chunks[current], date, line);
        host::emit(&mut chunks[current], "ecma:date", "getUTCFullYear", 1, line);
        chunks[current].emit_op(Op::F64_SUB, line);
        core_wasm::f64_const(&mut chunks[current], line, 12.0);
        chunks[current].emit_op(Op::F64_MUL, line);
    } else {
        get(&mut chunks[current], value, line);
        core_wasm::i32_const(&mut chunks[current], line, -1);
        ops::emit_dyn_add(&mut chunks[current], line);
        get(&mut chunks[current], date, line);
        host::emit(&mut chunks[current], "ecma:date", "getUTCMonth", 1, line);
        chunks[current].emit_op(Op::F64_SUB, line);
    }

    let delta = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], delta, line);
    get(&mut chunks[current], local_ms, line);
    get(&mut chunks[current], delta, line);
    datetime::emit_add_months(&mut chunks[current], DateTimePolicy::ISO, line);
    get(&mut chunks[current], offset, line);
    core_wasm::f64_const(&mut chunks[current], line, 1000.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    let ms = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], ms, line);
    prop_get(chunks, current, inst, "nano", line);
    let nano = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], nano, line);
    prop_get(chunks, current, inst, "kind", line);
    let kind = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], kind, line);
    make_time_from_ms_kind_slot(chunks, current, ms, nano, offset, kind, line);
}

pub fn emit_time_with_field(
    chunks: &mut [Chunk],
    current: usize,
    setter: &'static str,
    adjust_one_based: bool,
    line: u32,
) {
    let value = chunks[current].alloc_scratch(1);
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], inst, line);
    emit_epoch_milli_from_slot(chunks, current, inst, line);
    host::emit(&mut chunks[current], "ecma:date", "new", 1, line);
    let date = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], date, line);
    get(&mut chunks[current], date, line);
    get(&mut chunks[current], value, line);
    if adjust_one_based {
        core_wasm::i32_const(&mut chunks[current], line, -1);
        vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    }
    host::emit(&mut chunks[current], "ecma:date", setter, 2, line);
    let ms = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], ms, line);
    prop_get(chunks, current, inst, "nano", line);
    let nano = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], nano, line);
    prop_get(chunks, current, inst, "offset", line);
    let offset = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], offset, line);
    make_time_from_ms_slots(chunks, current, ms, nano, offset, "LocalDateTime", line);
}

pub fn emit_time_length_of_month(chunks: &mut [Chunk], current: usize, line: u32) {
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], inst, line);
    emit_epoch_milli_from_slot(chunks, current, inst, line);
    host::emit(&mut chunks[current], "ecma:date", "new", 1, line);
    let date = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], date, line);
    get(&mut chunks[current], date, line);
    host::emit(&mut chunks[current], "ecma:date", "getUTCFullYear", 1, line);
    get(&mut chunks[current], date, line);
    host::emit(&mut chunks[current], "ecma:date", "getUTCMonth", 1, line);
    datetime::emit_days_in_month(&mut chunks[current], MonthIndexing::ZeroBased, line);
}

pub fn emit_time_range_day(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    emit_time_length_of_month(chunks, current, line);
}

pub fn emit_time_is_leap_year(chunks: &mut [Chunk], current: usize, line: u32) {
    // Year extraction stays (the input is an Instant); the leap RULE is shared.
    // This previously round-tripped through `Date.UTC(y, 2, 0, …)` + getUTCDate
    // == 29 — the Feb-29 rollover, correct but two host calls and a temporary
    // Date to decide an integer question.
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], inst, line);
    emit_epoch_milli_from_slot(chunks, current, inst, line);
    host::emit(&mut chunks[current], "ecma:date", "new", 1, line);
    host::emit(&mut chunks[current], "ecma:date", "getUTCFullYear", 1, line);
    vybe_compiler::primitives::datetime::emit_is_leap_year(&mut chunks[current], line);
    // The shared rule answers a raw i32; the Java surface is a boolean —
    // without the box it PRINTED as `1`/`0`.
    ops::emit_i32_to_bool(&mut chunks[current], line);
}
pub fn emit_time_day_of_year(chunks: &mut [Chunk], current: usize, line: u32) {
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], inst, line);
    emit_epoch_milli_from_slot(chunks, current, inst, line);
    let ms = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], ms, line);
    get(&mut chunks[current], ms, line);
    get(&mut chunks[current], ms, line);
    host::emit(&mut chunks[current], "ecma:date", "new", 1, line);
    host::emit(&mut chunks[current], "ecma:date", "getUTCFullYear", 1, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    host::emit(&mut chunks[current], "ecma:date", "UTC", 7, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    core_wasm::f64_const(
        &mut chunks[current],
        line,
        vybe_compiler::primitives::datetime::MS_PER_DAY,
    );
    chunks[current].emit_op(Op::F64_DIV, line);
    host::emit(&mut chunks[current], "ecma:math", "floor", 1, line);
}

pub fn emit_time_day_of_week(chunks: &mut [Chunk], current: usize, line: u32) {
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], inst, line);
    emit_epoch_milli_from_slot(chunks, current, inst, line);
    host::emit(&mut chunks[current], "ecma:date", "new", 1, line);
    host::emit(&mut chunks[current], "ecma:date", "getUTCDay", 1, line);
    datetime::emit_weekday_in_base(&mut chunks[current], WeekdayBase::MondayOne, line);
}

/// `DayOfWeek.name` — the ISO number (Monday=1) spelled as Java's enum name.
pub fn emit_day_of_week_name(chunks: &mut [Chunk], current: usize, line: u32) {
    let n = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], n, line);
    let names = [
        "MONDAY",
        "TUESDAY",
        "WEDNESDAY",
        "THURSDAY",
        "FRIDAY",
        "SATURDAY",
    ];
    for (i, name) in names.iter().enumerate() {
        get(&mut chunks[current], n, line);
        core_wasm::f64_const(&mut chunks[current], line, (i + 1) as f64);
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_string_const(name, line);
        chunks[current].emit_else(line);
    }
    chunks[current].emit_string_const("SUNDAY", line);
    for _ in names {
        chunks[current].emit_end(line);
    }
}

/// `ZoneOffset.id` — the stored offset string, with Java's `Z` spelling
/// for UTC.
pub fn emit_zone_offset_id(chunks: &mut [Chunk], current: usize, line: u32) {
    let s = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], s, line);
    get(&mut chunks[current], s, line);
    chunks[current].emit_string_const("UTC", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("Z", line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], s, line);
    chunks[current].emit_end(line);
}

/// The epoch-ms of the THURSDAY of the value's ISO week — both week fields
/// derive from it (ISO 8601: a week belongs to the year holding its
/// Thursday).
fn emit_iso_thursday_ms(chunks: &mut [Chunk], current: usize, inst: u16, line: u32) {
    emit_epoch_milli_from_slot(chunks, current, inst, line);
    let ms = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], ms, line);
    get(&mut chunks[current], ms, line);
    core_wasm::f64_const(&mut chunks[current], line, 4.0);
    get(&mut chunks[current], ms, line);
    host::emit(&mut chunks[current], "ecma:date", "new", 1, line);
    host::emit(&mut chunks[current], "ecma:date", "getUTCDay", 1, line);
    datetime::emit_weekday_in_base(&mut chunks[current], WeekdayBase::MondayOne, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    core_wasm::f64_const(&mut chunks[current], line, 86_400_000.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
}

/// `temporal.with(field, value)` — same NAME-string fields as `get`.
pub fn emit_time_with(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let field = chunks[current].alloc_scratch(1);
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], field, line);
    set(&mut chunks[current], inst, line);

    for (name, setter) in [
        ("HOUR_OF_DAY", "setUTCHours"),
        ("DAY_OF_MONTH", "setUTCDate"),
    ] {
        get(&mut chunks[current], field, line);
        chunks[current].emit_string_const(name, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        get(&mut chunks[current], inst, line);
        get(&mut chunks[current], value, line);
        emit_time_with_field(chunks, current, setter, false, line);
        chunks[current].emit_else(line);
    }
    get(&mut chunks[current], inst, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// `temporal.get(field)` — the ChronoField/IsoFields constants arrive as
/// their NAME strings (that is what the tree registers the consts as).
pub fn emit_time_get(chunks: &mut [Chunk], current: usize, line: u32) {
    let field = chunks[current].alloc_scratch(1);
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], field, line);
    set(&mut chunks[current], inst, line);

    let arm = |chunks: &mut [Chunk], name: &str, line: u32| {
        get(&mut chunks[current], field, line);
        chunks[current].emit_string_const(name, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
    };

    arm(chunks, "DAY_OF_MONTH", line);
    get(&mut chunks[current], inst, line);
    emit_component(chunks, current, "getUTCDate", false, line);
    chunks[current].emit_else(line);

    arm(chunks, "HOUR_OF_DAY", line);
    get(&mut chunks[current], inst, line);
    emit_component(chunks, current, "getUTCHours", false, line);
    chunks[current].emit_else(line);

    arm(chunks, "WEEK_OF_WEEK_BASED_YEAR", line);
    {
        // trunc((thursday - jan1(year(thursday))) / week_ms) + 1
        emit_iso_thursday_ms(chunks, current, inst, line);
        let tms = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], tms, line);
        get(&mut chunks[current], tms, line);
        get(&mut chunks[current], tms, line);
        host::emit(&mut chunks[current], "ecma:date", "new", 1, line);
        host::emit(&mut chunks[current], "ecma:date", "getUTCFullYear", 1, line);
        for v in [0.0, 1.0, 0.0, 0.0, 0.0, 0.0] {
            core_wasm::f64_const(&mut chunks[current], line, v);
        }
        host::emit(&mut chunks[current], "ecma:date", "UTC", 7, line);
        chunks[current].emit_op(Op::F64_SUB, line);
        core_wasm::f64_const(&mut chunks[current], line, 604_800_000.0);
        chunks[current].emit_op(Op::F64_DIV, line);
        chunks[current].emit_op(Op::F64_TRUNC, line);
        core_wasm::f64_const(&mut chunks[current], line, 1.0);
        chunks[current].emit_op(Op::F64_ADD, line);
    }
    chunks[current].emit_else(line);

    arm(chunks, "WEEK_BASED_YEAR", line);
    emit_iso_thursday_ms(chunks, current, inst, line);
    host::emit(&mut chunks[current], "ecma:date", "new", 1, line);
    host::emit(&mut chunks[current], "ecma:date", "getUTCFullYear", 1, line);
    chunks[current].emit_else(line);

    // Unknown field: 0, never undefined.
    core_wasm::f64_const(&mut chunks[current], line, 0.0);
    for _ in 0..4 {
        chunks[current].emit_end(line);
    }
}

pub fn emit_truncated(chunks: &mut [Chunk], current: usize, line: u32) {
    let unit = chunks[current].alloc_scratch(1);
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], unit, line);
    set(&mut chunks[current], inst, line);
    prop_get(chunks, current, inst, "sec", line);
    let raw_sec = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], raw_sec, line);
    get(&mut chunks[current], unit, line);
    chunks[current].emit_string_const("HOURS", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], raw_sec, line);
    core_wasm::f64_const(&mut chunks[current], line, 3600.0);
    chunks[current].emit_op(Op::F64_DIV, line);
    host::emit(&mut chunks[current], "ecma:math", "floor", 1, line);
    core_wasm::f64_const(&mut chunks[current], line, 3600.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], unit, line);
    chunks[current].emit_string_const("MINUTES", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], raw_sec, line);
    core_wasm::f64_const(&mut chunks[current], line, 60.0);
    chunks[current].emit_op(Op::F64_DIV, line);
    host::emit(&mut chunks[current], "ecma:math", "floor", 1, line);
    core_wasm::f64_const(&mut chunks[current], line, 60.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], raw_sec, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    let sec = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], sec, line);
    get(&mut chunks[current], unit, line);
    chunks[current].emit_string_const("MILLIS", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    prop_get(chunks, current, inst, "nano", line);
    core_wasm::f64_const(&mut chunks[current], line, 1_000_000.0);
    chunks[current].emit_op(Op::F64_DIV, line);
    host::emit(&mut chunks[current], "ecma:math", "floor", 1, line);
    core_wasm::f64_const(&mut chunks[current], line, 1_000_000.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_end(line);
    let nano = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], nano, line);
    make_instant_from_slots(chunks, current, sec, nano, line);
}

pub fn emit_hash_code(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_to_epoch_milli(chunks, current, line);
}
