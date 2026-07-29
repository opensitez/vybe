//! Small Java `java.time` adapter backed by ECMA date helpers.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;
use vybe_compiler::primitives::instructions::{core_wasm, host};

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

pub fn emit_local_time_parse(chunks: &mut [Chunk], current: usize, line: u32) {
    let time = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], time, line);
    chunks[current].emit_string_const("1970-01-01T", line);
    get(&mut chunks[current], time, line);
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
    prop_get(chunks, current, out, "offset", line);
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

pub fn emit_duration_to_hours(chunks: &mut [Chunk], current: usize, line: u32) {
    core_wasm::f64_const(&mut chunks[current], line, 3600.0);
    chunks[current].emit_op(Op::F64_DIV, line);
}

pub fn emit_duration_to_minutes(chunks: &mut [Chunk], current: usize, line: u32) {
    core_wasm::f64_const(&mut chunks[current], line, 60.0);
    chunks[current].emit_op(Op::F64_DIV, line);
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

pub fn emit_period_days(_chunks: &mut [Chunk], _current: usize, _line: u32) {}

pub fn emit_period_months(_chunks: &mut [Chunk], _current: usize, _line: u32) {}

pub fn emit_period_between(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_duration_between(chunks, current, line);
    core_wasm::f64_const(&mut chunks[current], line, 86400.0);
    chunks[current].emit_op(Op::F64_DIV, line);
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

pub fn emit_local_date_string(chunks: &mut [Chunk], current: usize, line: u32) {
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
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 10);
    host::emit(&mut chunks[current], "ecma:string", "substring", 3, line);
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
    host::emit(&mut chunks[current], "ecma:date", "new", 1, line);
    let date = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], date, line);
    get(&mut chunks[current], date, line);
    host::emit(&mut chunks[current], "ecma:date", "getUTCMonth", 1, line);
    get(&mut chunks[current], delta, line);
    core_wasm::f64_const(&mut chunks[current], line, sign);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    let new_month = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], new_month, line);
    get(&mut chunks[current], date, line);
    get(&mut chunks[current], new_month, line);
    host::emit(&mut chunks[current], "ecma:date", "setUTCMonth", 2, line);
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
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    host::emit(&mut chunks[current], "ecma:date", "UTC", 7, line);
    host::emit(&mut chunks[current], "ecma:date", "new", 1, line);
    host::emit(&mut chunks[current], "ecma:date", "getUTCDate", 1, line);
}

pub fn emit_time_range_day(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    emit_time_length_of_month(chunks, current, line);
}

pub fn emit_time_is_leap_year(chunks: &mut [Chunk], current: usize, line: u32) {
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], inst, line);
    emit_epoch_milli_from_slot(chunks, current, inst, line);
    host::emit(&mut chunks[current], "ecma:date", "new", 1, line);
    host::emit(&mut chunks[current], "ecma:date", "getUTCFullYear", 1, line);
    let year = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], year, line);
    get(&mut chunks[current], year, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    host::emit(&mut chunks[current], "ecma:date", "UTC", 7, line);
    host::emit(&mut chunks[current], "ecma:date", "new", 1, line);
    host::emit(&mut chunks[current], "ecma:date", "getUTCDate", 1, line);
    core_wasm::i32_const(&mut chunks[current], line, 29);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
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
    core_wasm::f64_const(&mut chunks[current], line, 86_400_000.0);
    chunks[current].emit_op(Op::F64_DIV, line);
    host::emit(&mut chunks[current], "ecma:math", "floor", 1, line);
}

pub fn emit_time_day_of_week(chunks: &mut [Chunk], current: usize, line: u32) {
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], inst, line);
    emit_epoch_milli_from_slot(chunks, current, inst, line);
    host::emit(&mut chunks[current], "ecma:date", "new", 1, line);
    host::emit(&mut chunks[current], "ecma:date", "getUTCDay", 1, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("MONDAY", line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("SUNDAY", line);
    chunks[current].emit_end(line);
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
