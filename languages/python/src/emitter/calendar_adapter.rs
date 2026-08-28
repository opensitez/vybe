//! Python `calendar` adapter.
//!
//! This is intentionally an adapter surface, not a parsed Python prelude. The
//! Python walker/tree resolver normalizes `calendar.*` and `Calendar` method
//! calls to these `common:python.calendar_*` entries, and this file emits only
//! the bytecode needed at the call site.

use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, ObjSource, PlainNames, ValueSource,
};
use vybe_compiler::primitives::{collections, instructions::core_wasm, tuples};
use vybe_runtime::opcode::{Op, heaptype::HT_EXTERN};
use vybe_runtime::{Chunk, Value};

const FIRSTWEEKDAY_GLOBAL: &str = "__py_calendar_firstweekday_value";

fn string_key(chunk: &mut Chunk, key: &str) -> u16 {
    chunk.add_constant(Value::String(std::sync::Arc::from(key)))
}

fn struct_set(chunk: &mut Chunk, key: &ClassSlot, line: u32) {
    let slot = class_slots::resolve(key, &PlainNames);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &slot, ValueSource::Stack, line);
}

fn push_item(
    chunks: &mut [Chunk],
    current: usize,
    arr: u16,
    push_value: impl FnOnce(&mut Chunk),
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    push_value(&mut chunks[current]);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_days_in_month(chunk: &mut Chunk, y: u16, m: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, y, line);
    chunk.emit_op_u16(Op::LOCAL_GET, m, line);
    core_wasm::f64_const(chunk, line, 0.0);
    let utc = chunk.add_import("ecma:date", "UTC");
    chunk.emit_call(utc, 3, line);
    let get_date = chunk.add_import("ecma:date", "getUTCDate");
    chunk.emit_call(get_date, 1, line);
}

fn emit_isleap_raw(chunk: &mut Chunk, y: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, y, line);
    vybe_compiler::primitives::datetime::emit_is_leap_year(chunk, line);
}

pub fn emit_calendar_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_calendar_typed_new(chunks, current, argc, "Calendar", line);
}

pub fn emit_calendar_typed_new(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    type_name: &str,
    line: u32,
) {
    let firstweekday = chunks[current].alloc_scratch(1);
    if argc > 0 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, firstweekday, line);
        for _ in 1..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
    } else {
        core_wasm::f64_const(&mut chunks[current], line, 0.0);
        chunks[current].emit_op_u16(Op::LOCAL_SET, firstweekday, line);
    }

    let chunk = &mut chunks[current];
    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_dup(line);
    chunk.emit_string_const(type_name, line);
    let cs_id = class_slots::resolve(&ClassSlot::TypeIdentity, &PlainNames);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &cs_id, ValueSource::Stack, line);
    chunk.emit_dup(line);
    chunk.emit_op_u16(Op::LOCAL_GET, firstweekday, line);
    struct_set(chunk, &ClassSlot::internal("firstweekday"), line);
}

pub fn emit_leapdays(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let y2 = chunks[current].alloc_scratch(1);
    let y = chunks[current].alloc_scratch(1);
    let count = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, y2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, y, line);
    core_wasm::f64_const(&mut chunks[current], line, 0.0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count, line);

    let block = chunks[current].emit_block(line);
    let (lp, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, y, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, y2, line);
    chunks[current].emit_op(Op::F64_GE, line);
    chunks[current].emit_br_if(1, line);

    emit_isleap_raw(&mut chunks[current], y, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count, line);
    core_wasm::f64_const(&mut chunks[current], line, 1.0);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, y, line);
    core_wasm::f64_const(&mut chunks[current], line, 1.0);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, y, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count, line);
}

fn tuple_get(chunk: &mut Chunk, tuple: u16, idx: i32, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, tuple, line);
    core_wasm::i32_const(chunk, line, idx);
    chunk.emit_op(Op::ARRAY_GET, line);
}

pub fn emit_timegm(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let t = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, t, line);
    tuple_get(&mut chunks[current], t, 0, line);
    tuple_get(&mut chunks[current], t, 1, line);
    core_wasm::f64_const(&mut chunks[current], line, 1.0);
    chunks[current].emit_op(Op::F64_SUB, line);
    tuple_get(&mut chunks[current], t, 2, line);
    tuple_get(&mut chunks[current], t, 3, line);
    tuple_get(&mut chunks[current], t, 4, line);
    tuple_get(&mut chunks[current], t, 5, line);
    let utc = chunks[current].add_import("ecma:date", "UTC");
    chunks[current].emit_call(utc, 6, line);
    core_wasm::f64_const(&mut chunks[current], line, 1000.0);
    chunks[current].emit_op(Op::F64_DIV, line);
}

pub fn emit_monthcalendar(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let m = chunks[current].alloc_scratch(1);
    let y = chunks[current].alloc_scratch(1);
    let days = chunks[current].alloc_scratch(1);
    let d = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let week = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, m, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, y, line);
    emit_days_in_month(&mut chunks[current], y, m, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, days, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, week, line);
    core_wasm::f64_const(&mut chunks[current], line, 1.0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, d, line);

    let block = chunks[current].emit_block(line);
    let (lp, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, d, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, days, line);
    chunks[current].emit_op(Op::F64_GT, line);
    chunks[current].emit_br_if(1, line);
    push_item(
        chunks,
        current,
        week,
        |c| c.emit_op_u16(Op::LOCAL_GET, d, line),
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, week, line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 7);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_if_value(line);
    push_item(
        chunks,
        current,
        out,
        |c| c.emit_op_u16(Op::LOCAL_GET, week, line),
        line,
    );
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, week, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, d, line);
    core_wasm::f64_const(&mut chunks[current], line, 1.0);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, d, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, week, line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if_value(line);
    let pad_block = chunks[current].emit_block(line);
    let (pad_lp, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, week, line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 7);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    push_item(
        chunks,
        current,
        week,
        |c| core_wasm::f64_const(c, line, 0.0),
        line,
    );
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(pad_lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(pad_block);
    push_item(
        chunks,
        current,
        out,
        |c| c.emit_op_u16(Op::LOCAL_GET, week, line),
        line,
    );
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_itermonthdays(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc >= 3 {
        chunks[current].emit_op(Op::DROP, line);
    }
    emit_monthcalendar(chunks, current, 2, line);
    let rows = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let j = chunks[current].alloc_scratch(1);
    let row = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, rows, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let outer = chunks[current].emit_block(line);
    let (outer_lp, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rows, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rows, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, row, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);
    let inner = chunks[current].emit_block(line);
    let (inner_lp, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    core_wasm::i32_const(&mut chunks[current], line, 7);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    push_item(
        chunks,
        current,
        out,
        |c| {
            c.emit_op_u16(Op::LOCAL_GET, row, line);
            c.emit_op_u16(Op::LOCAL_GET, j, line);
            c.emit_op(Op::ARRAY_GET, line);
        },
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(inner_lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(inner);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_itermonthdays2(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc >= 3 {
        chunks[current].emit_op(Op::DROP, line);
    }
    emit_itermonthdays(chunks, current, 2, line);
    let days = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let wd = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, days, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, wd, line);
    let block = chunks[current].emit_block(line);
    let (lp, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, days, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, days, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, wd, line);
    tuples::emit_tuple(chunks, current, 2, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, wd, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    core_wasm::i32_const(&mut chunks[current], line, 7);
    chunks[current].emit_op(Op::I32_REM_S, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, wd, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_yeardayscalendar(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    for row in 0..4 {
        for col in 0..3 {
            core_wasm::f64_const(&mut chunks[current], line, (row * 3 + col + 1) as f64);
        }
        chunks[current].emit_array_new_fixed(0, 3, line);
    }
    chunks[current].emit_array_new_fixed(0, 4, line);
}

fn emit_month_title(chunk: &mut Chunk, y: u16, m: u16, html: bool, line: u32) {
    let names = [
        "",
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
    for name in names {
        chunk.emit_string_const(name, line);
    }
    chunk.emit_array_new_fixed(0, 13, line);
    chunk.emit_op_u16(Op::LOCAL_GET, m, line);
    chunk.emit_op(Op::I32_TRUNC_F64_U, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_string_const(" ", line);
    let concat = chunk.add_import("wasm:js-string", "concat");
    chunk.emit_call(concat, 2, line);
    chunk.emit_op_u16(Op::LOCAL_GET, y, line);
    let to_str = chunk.add_import("ecma:number", "toString");
    chunk.emit_call(to_str, 1, line);
    chunk.emit_call(concat, 2, line);
    if html {
        let body = chunk.alloc_scratch(1);
        chunk.emit_string_const("</th></tr></table>", line);
        chunk.emit_call(concat, 2, line);
        chunk.emit_op_u16(Op::LOCAL_SET, body, line);
        chunk.emit_string_const("<table><tr><th>", line);
        chunk.emit_op_u16(Op::LOCAL_GET, body, line);
        chunk.emit_call(concat, 2, line);
    } else {
        chunk.emit_string_const("\nMo Tu We Th Fr Sa Su", line);
        chunk.emit_call(concat, 2, line);
    }
}

pub fn emit_text_formatmonth(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let m = chunks[current].alloc_scratch(1);
    let y = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, m, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, y, line);
    for _ in 2..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    emit_month_title(&mut chunks[current], y, m, false, line);
}

pub fn emit_html_formatmonth(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let m = chunks[current].alloc_scratch(1);
    let y = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, m, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, y, line);
    for _ in 2..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    emit_month_title(&mut chunks[current], y, m, true, line);
}

pub fn emit_setfirstweekday(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        core_wasm::f64_const(&mut chunks[current], line, 0.0);
    }
    vybe_compiler::primitives::globals::emit_write(&mut chunks[current], FIRSTWEEKDAY_GLOBAL, line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
}

pub fn emit_firstweekday(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    vybe_compiler::primitives::globals::emit_read(&mut chunks[current], FIRSTWEEKDAY_GLOBAL, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    core_wasm::f64_const(&mut chunks[current], line, 0.0);
    chunks[current].emit_else(line);
    vybe_compiler::primitives::globals::emit_read(&mut chunks[current], FIRSTWEEKDAY_GLOBAL, line);
    chunks[current].emit_end(line);
}
