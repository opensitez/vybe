//! Dart core library adapters for Duration, DateTime, and Uri.

use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};
use vybe_emitter::collections;
use vybe_emitter::instructions::{core_wasm, host};

fn key(chunk: &mut Chunk, name: &str) -> u16 {
    chunk.add_constant(Value::String(Arc::from(name)))
}

fn set_field(chunk: &mut Chunk, name: &str, line: u32) {
    let k = key(chunk, name);
    chunk.emit_op_u16(Op::STRUCT_SET, k, line);
    chunk.emit_op(Op::DROP, line);
}

fn get_field(chunk: &mut Chunk, name: &str, line: u32) {
    let k = key(chunk, name);
    chunk.emit_op_u16(Op::STRUCT_GET, k, line);
}

fn set_string(chunk: &mut Chunk, name: &str, value: &str, line: u32) {
    core_wasm::dup(chunk, line);
    chunk.emit_string_const(value, line);
    set_field(chunk, name, line);
}

fn set_bool(chunk: &mut Chunk, name: &str, value: bool, line: u32) {
    core_wasm::dup(chunk, line);
    chunk.emit_bool_const(value, line);
    set_field(chunk, name, line);
}

fn obj_new(chunk: &mut Chunk, line: u32) {
    host::emit(chunk, "ecma:object", "new", 0, line);
}

fn date_get(chunks: &mut [Chunk], current: usize, ms_slot: u16, getter: &'static str, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    host::emit(&mut chunks[current], "ecma:date", getter, 1, line);
}

fn wrap_duration_ms(chunk: &mut Chunk, line: u32) {
    let ms = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, ms, line);
    obj_new(chunk, line);
    set_string(chunk, "__type", "Duration", line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    set_field(chunk, "inMilliseconds", line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    chunk.emit_f64_const(1000.0, line);
    chunk.emit_op(Op::F64_DIV, line);
    set_field(chunk, "inSeconds", line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    chunk.emit_f64_const(60_000.0, line);
    chunk.emit_op(Op::F64_DIV, line);
    set_field(chunk, "inMinutes", line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    chunk.emit_f64_const(3_600_000.0, line);
    chunk.emit_op(Op::F64_DIV, line);
    set_field(chunk, "inHours", line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    chunk.emit_f64_const(86_400_000.0, line);
    chunk.emit_op(Op::F64_DIV, line);
    set_field(chunk, "inDays", line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    chunk.emit_f64_const(0.0, line);
    vybe_emitter::ops::emit_dyn_lt(chunk, line);
    vybe_emitter::ops::emit_i32_to_bool(chunk, line);
    set_field(chunk, "isNegative", line);
}

fn duration_ms_from_obj(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    get_field(chunk, "inMilliseconds", line);
}

pub fn emit_duration_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_f64_const(0.0, line);
    }
    wrap_duration_ms(&mut chunks[current], line);
}

pub fn emit_duration_zero(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_f64_const(0.0, line);
    wrap_duration_ms(&mut chunks[current], line);
}

pub fn emit_duration_abs(chunks: &mut [Chunk], current: usize, line: u32) {
    let slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
    duration_ms_from_obj(&mut chunks[current], slot, line);
    chunks[current].emit_op(Op::F64_ABS, line);
    wrap_duration_ms(&mut chunks[current], line);
}

pub fn emit_duration_negate(chunks: &mut [Chunk], current: usize, line: u32) {
    let slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
    duration_ms_from_obj(&mut chunks[current], slot, line);
    chunks[current].emit_f64_const(-1.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    wrap_duration_ms(&mut chunks[current], line);
}

fn utc_from_stack(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = chunks[current].alloc_scratch(6);
    let second = base;
    let minute = base + 1;
    let hour = base + 2;
    let day = base + 3;
    let month = base + 4;
    let year = base + 5;
    if argc >= 6 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, second, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, minute, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, hour, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, day, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, month, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, year, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, year, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, month, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, day, line);
    if argc >= 6 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, hour, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, minute, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, second, line);
    } else {
        chunks[current].emit_f64_const(0.0, line);
        chunks[current].emit_f64_const(0.0, line);
        chunks[current].emit_f64_const(0.0, line);
    }
    host::emit(&mut chunks[current], "ecma:date", "UTC", 6, line);
}

fn wrap_datetime_ms(chunks: &mut [Chunk], current: usize, is_utc: bool, line: u32) {
    let ms = chunks[current].alloc_scratch(2);
    let dow = ms + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, ms, line);
    obj_new(&mut chunks[current], line);
    set_string(&mut chunks[current], "__type", "DateTime", line);
    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ms, line);
    set_field(&mut chunks[current], "millisecondsSinceEpoch", line);
    for (field, getter, add_one) in [
        ("year", "getUTCFullYear", false),
        ("month", "getUTCMonth", true),
        ("day", "getUTCDate", false),
        ("hour", "getUTCHours", false),
        ("minute", "getUTCMinutes", false),
        ("second", "getUTCSeconds", false),
    ] {
        core_wasm::dup(&mut chunks[current], line);
        date_get(chunks, current, ms, getter, line);
        if add_one {
            chunks[current].emit_f64_const(1.0, line);
            chunks[current].emit_op(Op::F64_ADD, line);
        }
        set_field(&mut chunks[current], field, line);
    }
    date_get(chunks, current, ms, "getUTCDay", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dow, line);
    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dow, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dow, line);
    chunks[current].emit_else(line);
    chunks[current].emit_f64_const(7.0, line);
    chunks[current].emit_end(line);
    set_field(&mut chunks[current], "weekday", line);
    set_bool(&mut chunks[current], "isUtc", is_utc, line);
}

fn datetime_ms_from_obj(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    get_field(chunk, "millisecondsSinceEpoch", line);
}

fn comparable_value_from_obj(chunk: &mut Chunk, slot: u16, out: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    get_field(chunk, "millisecondsSinceEpoch", line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    host::emit(chunk, "wasm:js-undefined", "test", 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    get_field(chunk, "inMilliseconds", line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);
    chunk.emit_end(line);
}

pub fn emit_datetime_new(chunks: &mut [Chunk], current: usize, argc: u8, is_utc: bool, line: u32) {
    utc_from_stack(chunks, current, argc, line);
    wrap_datetime_ms(chunks, current, is_utc, line);
}

pub fn emit_datetime_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let dur = chunks[current].alloc_scratch(2);
    let dt = dur + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, dur, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dt, line);
    datetime_ms_from_obj(&mut chunks[current], dt, line);
    duration_ms_from_obj(&mut chunks[current], dur, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    wrap_datetime_ms(chunks, current, false, line);
}

pub fn emit_dart_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(2);
    let receiver = value + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver, line);
    get_field(&mut chunks[current], "__type", line);
    chunks[current].emit_string_const("DateTime", line);
    chunks[current].emit_op(Op::EQ, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    emit_datetime_add(chunks, current, line);

    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    collections::emit_push(chunks, current, line);

    chunks[current].emit_end(line);
}

pub fn emit_datetime_subtract(chunks: &mut [Chunk], current: usize, line: u32) {
    let dur = chunks[current].alloc_scratch(2);
    let dt = dur + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, dur, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dt, line);
    datetime_ms_from_obj(&mut chunks[current], dt, line);
    duration_ms_from_obj(&mut chunks[current], dur, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    wrap_datetime_ms(chunks, current, false, line);
}

pub fn emit_datetime_difference(chunks: &mut [Chunk], current: usize, line: u32) {
    let right = chunks[current].alloc_scratch(2);
    let left = right + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, right, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left, line);
    datetime_ms_from_obj(&mut chunks[current], left, line);
    datetime_ms_from_obj(&mut chunks[current], right, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    wrap_duration_ms(&mut chunks[current], line);
}

fn compare_ms(chunks: &mut [Chunk], current: usize, line: u32, op: Op) {
    let right = chunks[current].alloc_scratch(2);
    let left = right + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, right, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left, line);
    datetime_ms_from_obj(&mut chunks[current], left, line);
    datetime_ms_from_obj(&mut chunks[current], right, line);
    chunks[current].emit_op(op, line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_datetime_is_before(chunks: &mut [Chunk], current: usize, line: u32) {
    compare_ms(chunks, current, line, Op::F64_LT);
}

pub fn emit_datetime_is_after(chunks: &mut [Chunk], current: usize, line: u32) {
    compare_ms(chunks, current, line, Op::F64_GT);
}

pub fn emit_datetime_same_moment(chunks: &mut [Chunk], current: usize, line: u32) {
    compare_ms(chunks, current, line, Op::F64_EQ);
}

pub fn emit_compare_to(chunks: &mut [Chunk], current: usize, line: u32) {
    let right = chunks[current].alloc_scratch(4);
    let left = right + 1;
    let r = right + 2;
    let l = right + 3;
    chunks[current].emit_op_u16(Op::LOCAL_SET, right, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    host::emit(&mut chunks[current], "wasm:js-string", "test", 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right, line);
    host::emit(&mut chunks[current], "wasm:js-string", "compare", 2, line);
    chunks[current].emit_else(line);
    comparable_value_from_obj(&mut chunks[current], left, l, line);
    comparable_value_from_obj(&mut chunks[current], right, r, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, l, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, r, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, l, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, r, line);
    vybe_emitter::ops::emit_dyn_gt(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn get_url_prop(chunk: &mut Chunk, url_slot: u16, prop: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, url_slot, line);
    get_field(chunk, prop, line);
}

fn wrap_url(chunks: &mut [Chunk], current: usize, line: u32) {
    let url = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, url, line);
    obj_new(&mut chunks[current], line);
    set_string(&mut chunks[current], "__type", "Uri", line);
    core_wasm::dup(&mut chunks[current], line);
    get_url_prop(&mut chunks[current], url, "protocol", line);
    chunks[current].emit_string_const(":", line);
    chunks[current].emit_string_const("", line);
    host::emit(&mut chunks[current], "ecma:string", "replace", 3, line);
    set_field(&mut chunks[current], "scheme", line);
    for (dart, web) in [
        ("host", "hostname"),
        ("authority", "host"),
        ("path", "pathname"),
        ("queryParameters", "searchParams"),
        ("origin", "origin"),
    ] {
        core_wasm::dup(&mut chunks[current], line);
        get_url_prop(&mut chunks[current], url, web, line);
        set_field(&mut chunks[current], dart, line);
    }
    core_wasm::dup(&mut chunks[current], line);
    get_url_prop(&mut chunks[current], url, "port", line);
    set_field(&mut chunks[current], "port", line);
    core_wasm::dup(&mut chunks[current], line);
    get_url_prop(&mut chunks[current], url, "search", line);
    chunks[current].emit_string_const("?", line);
    chunks[current].emit_string_const("", line);
    host::emit(&mut chunks[current], "ecma:string", "replace", 3, line);
    set_field(&mut chunks[current], "query", line);
    core_wasm::dup(&mut chunks[current], line);
    get_url_prop(&mut chunks[current], url, "hash", line);
    chunks[current].emit_string_const("#", line);
    chunks[current].emit_string_const("", line);
    host::emit(&mut chunks[current], "ecma:string", "replace", 3, line);
    set_field(&mut chunks[current], "fragment", line);
    core_wasm::dup(&mut chunks[current], line);
    get_url_prop(&mut chunks[current], url, "href", line);
    set_field(&mut chunks[current], "href", line);
    set_bool(&mut chunks[current], "hasScheme", true, line);
    set_bool(&mut chunks[current], "hasAuthority", true, line);
    set_bool(&mut chunks[current], "isAbsolute", true, line);
}

pub fn emit_uri_parse(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "web:url", "parse", 1, line);
    wrap_url(chunks, current, line);
}

pub fn emit_uri_http(chunks: &mut [Chunk], current: usize, argc: u8, https: bool, line: u32) {
    if argc > 2 {
        chunks[current].emit_op(Op::DROP, line);
    }
    if argc > 3 {
        chunks[current].emit_op(Op::DROP, line);
    }
    let path = chunks[current].alloc_scratch(2);
    let host_slot = path + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, path, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, host_slot, line);
    chunks[current].emit_string_const(if https { "https://" } else { "http://" }, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, host_slot, line);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    emit_uri_parse(chunks, current, line);
}

pub fn emit_uri_file(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_string_const("file://", line);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    emit_uri_parse(chunks, current, line);
}

pub fn emit_uri_normalize_path(_chunks: &mut [Chunk], _current: usize, _line: u32) {}
pub fn emit_uri_replace(_chunks: &mut [Chunk], _current: usize, _line: u32) {}

pub fn emit_uri_resolve(chunks: &mut [Chunk], current: usize, line: u32) {
    let rel = chunks[current].alloc_scratch(2);
    let base = rel + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, rel, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rel, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    get_field(&mut chunks[current], "href", line);
    host::emit(&mut chunks[current], "web:url", "parse", 2, line);
    wrap_url(chunks, current, line);
}

pub fn emit_uri_resolve_uri(chunks: &mut [Chunk], current: usize, line: u32) {
    let rel = chunks[current].alloc_scratch(2);
    let base = rel + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, rel, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rel, line);
    get_field(&mut chunks[current], "href", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    get_field(&mut chunks[current], "href", line);
    host::emit(&mut chunks[current], "web:url", "parse", 2, line);
    wrap_url(chunks, current, line);
}

pub fn emit_list_filled(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(2);
    let length = value + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, length, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, length, line);
    collections::emit_new_with_length(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, length, line);
    collections::emit_fill(chunks, current, line);
}
