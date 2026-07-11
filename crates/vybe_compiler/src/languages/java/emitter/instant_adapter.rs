//! Small Java `java.time` adapter backed by ECMA date helpers.

use crate::emitter::instructions::{core_wasm, host};
use vybe_bytecode::opcode::Op;
use vybe_bytecode::Chunk;

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
    crate::emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_else(line);
    emit_epoch_milli_from_slot(chunks, current, inst, line);
    emit_epoch_milli_from_slot(chunks, current, other, line);
    crate::emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
        crate::emitter::ops::emit_dyn_gt(&mut chunks[current], line);
    } else {
        crate::emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    }
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    crate::emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_compare(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    crate::emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    crate::emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
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

pub fn emit_duration_between(chunks: &mut [Chunk], current: usize, line: u32) {
    let end = chunks[current].alloc_scratch(1);
    let start = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], end, line);
    set(&mut chunks[current], start, line);
    prop_get(chunks, current, end, "sec", line);
    prop_get(chunks, current, start, "sec", line);
    chunks[current].emit_op(Op::F64_SUB, line);
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
    core_wasm::f64_const(&mut chunks[current], line, 3600.0);
    chunks[current].emit_op(Op::F64_MUL, line);
}

pub fn emit_zone_id_utc(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
}

pub fn emit_get_offset(chunks: &mut [Chunk], current: usize, line: u32) {
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], inst, line);
    prop_get(chunks, current, inst, "offset", line);
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
    core_wasm::f64_const(&mut chunks[current], line, 1000.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    host::emit(&mut chunks[current], "ecma:date", "new", 1, line);
    host::emit(&mut chunks[current], "ecma:date", method, 1, line);
    if add_one {
        core_wasm::i32_const(&mut chunks[current], line, 1);
        crate::emitter::ops::emit_dyn_add(&mut chunks[current], line);
    }
}

pub fn emit_local_date_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], inst, line);
    emit_epoch_milli_from_slot(chunks, current, inst, line);
    prop_get(chunks, current, inst, "offset", line);
    core_wasm::f64_const(&mut chunks[current], line, 1000.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    host::emit(&mut chunks[current], "ecma:date", "new", 1, line);
    host::emit(&mut chunks[current], "ecma:date", "toISOString", 1, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 10);
    host::emit(&mut chunks[current], "ecma:string", "substring", 3, line);
}

pub fn emit_truncated(chunks: &mut [Chunk], current: usize, line: u32) {
    let unit = chunks[current].alloc_scratch(1);
    let inst = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], unit, line);
    set(&mut chunks[current], inst, line);
    prop_get(chunks, current, inst, "sec", line);
    let sec = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], sec, line);
    get(&mut chunks[current], unit, line);
    chunks[current].emit_string_const("SECONDS", line);
    crate::emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_else(line);
    prop_get(chunks, current, inst, "nano", line);
    core_wasm::f64_const(&mut chunks[current], line, 1_000_000.0);
    chunks[current].emit_op(Op::F64_DIV, line);
    host::emit(&mut chunks[current], "ecma:math", "floor", 1, line);
    core_wasm::f64_const(&mut chunks[current], line, 1_000_000.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_end(line);
    let nano = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], nano, line);
    make_instant_from_slots(chunks, current, sec, nano, line);
}

pub fn emit_hash_code(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_to_epoch_milli(chunks, current, line);
}
