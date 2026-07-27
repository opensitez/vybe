//! Python `time` adapter — bytecode-only.
//!
//! Composes `ecma:date.*`. The centrepiece is `struct_time`, which is both
//! indexable (`t[0]`, `t[0:2]`, unpacking) AND field-named (`t.tm_year`) — i.e.
//! a **named tuple**. Built through the shared `tuples::emit_named_tuple`, so
//! it indexes/slices/unpacks/reprs (`time.struct_time(...)`) with no bespoke
//! runtime type. No new host fns.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;
use vybe_compiler::compiler::instructions::core_wasm;

const MS_PER_SECOND: f64 = 1_000.0;
const MS_PER_DAY: f64 = 86_400_000.0;

/// `struct_time`'s nine fields, in CPython order.
const TM_FIELDS: &[&str] = &[
    "tm_year", "tm_mon", "tm_mday", "tm_hour", "tm_min", "tm_sec", "tm_wday", "tm_yday", "tm_isdst",
];

fn getter(chunk: &mut Chunk, ms: u16, name: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    let idx = chunk.add_import("ecma:date", name);
    chunk.emit_call(idx, 1, line);
}

/// Build a `struct_time` from a millisecond timestamp already in `ms`.
/// Stack: `[]` → `[struct_time]`.
fn emit_struct_time_from_ms(chunks: &mut [Chunk], current: usize, ms: u16, line: u32) {
    let chunk = &mut chunks[current];
    // Push the nine fields in order. `getUTCMonth` is 0-based → +1; `getUTCDay`
    // is Sunday=0 → Python's Monday=0 via `(d + 6) % 7`.
    getter(chunk, ms, "getUTCFullYear", line);
    getter(chunk, ms, "getUTCMonth", line);
    core_wasm::f64_const(chunk, line, 1.0);
    chunk.emit_op(Op::F64_ADD, line);
    getter(chunk, ms, "getUTCDate", line);
    getter(chunk, ms, "getUTCHours", line);
    getter(chunk, ms, "getUTCMinutes", line);
    getter(chunk, ms, "getUTCSeconds", line);
    getter(chunk, ms, "getUTCDay", line);
    core_wasm::i32_const(chunk, line, 6);
    chunk.emit_op(Op::I32_ADD, line);
    core_wasm::i32_const(chunk, line, 7);
    chunk.emit_op(Op::I32_REM_S, line);

    // tm_yday = whole days from Jan 1 (UTC) of this year + 1.
    let year_start = chunk.alloc_scratch(1);
    getter(chunk, ms, "getUTCFullYear", line);
    core_wasm::f64_const(chunk, line, 0.0); // month 0
    core_wasm::f64_const(chunk, line, 1.0); // day 1
    let utc = chunk.add_import("ecma:date", "UTC");
    chunk.emit_call(utc, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, year_start, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    chunk.emit_op_u16(Op::LOCAL_GET, year_start, line);
    chunk.emit_op(Op::F64_SUB, line);
    core_wasm::f64_const(chunk, line, MS_PER_DAY);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    core_wasm::f64_const(chunk, line, 1.0);
    chunk.emit_op(Op::F64_ADD, line);

    // tm_isdst = 0 (UTC, no DST).
    core_wasm::i32_const(chunk, line, 0);

    emit_pack_and_name(chunks, current, line);
}

/// The nine field values are on the stack — pack them and stamp the named-tuple
/// metadata (tag + `tm_*` keys + `__fields` + `__typename`). Stack: `[v0..v8]` → `[struct_time]`.
fn emit_pack_and_name(chunks: &mut [Chunk], current: usize, line: u32) {
    let n = TM_FIELDS.len() as u16;
    let base = chunks[current].alloc_scratch(n);
    vybe_compiler::compiler::collections::emit_pack_n(chunks, current, n, base, line);
    let names: Vec<Option<String>> = TM_FIELDS.iter().map(|f| Some((*f).to_string())).collect();
    vybe_compiler::compiler::tuples::emit_named_tuple(chunks, current, &names, Some("struct_time"), line);
}

/// `time.gmtime(secs=now)` / `time.localtime(secs=now)` — UTC struct_time.
/// Stack: `[secs?]` → `[struct_time]`.
pub fn emit_gmtime(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let ms = chunks[current].alloc_scratch(1);
    if argc >= 1 {
        core_wasm::f64_const(&mut chunks[current], line, MS_PER_SECOND);
        chunks[current].emit_op(Op::F64_MUL, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, ms, line);
    } else {
        let now = chunks[current].add_import("ecma:date", "now");
        chunks[current].emit_call(now, 0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, ms, line);
    }
    emit_struct_time_from_ms(chunks, current, ms, line);
}

/// `time.struct_time(seq)` — wrap a 9-element sequence as a struct_time.
/// Stack: `[seq]` → `[struct_time]`.
pub fn emit_struct_time(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let seq = chunks[current].alloc_scratch(1);
    // Copy so the struct_time doesn't alias the caller's list.
    let from = chunks[current].add_import("ecma:array", "from");
    chunks[current].emit_call(from, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, seq, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, seq, line);
    let names: Vec<Option<String>> = TM_FIELDS.iter().map(|f| Some((*f).to_string())).collect();
    vybe_compiler::compiler::tuples::emit_named_tuple(chunks, current, &names, Some("struct_time"), line);
}

/// `time.mktime(t)` — struct_time → seconds since epoch. Stack: `[t]` → `[num]`.
pub fn emit_mktime(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let t = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, t, line);
    let chunk = &mut chunks[current];
    // UTC(year, mon-1, mday, hour, min, sec) / 1000.
    let read = |chunk: &mut Chunk, i: i32| {
        chunk.emit_op_u16(Op::LOCAL_GET, t, line);
        core_wasm::i32_const(chunk, line, i);
        chunk.emit_op(Op::ARRAY_GET, line);
    };
    read(chunk, 0);
    read(chunk, 1);
    core_wasm::f64_const(chunk, line, 1.0);
    chunk.emit_op(Op::F64_SUB, line);
    read(chunk, 2);
    read(chunk, 3);
    read(chunk, 4);
    read(chunk, 5);
    let utc = chunk.add_import("ecma:date", "UTC");
    chunk.emit_call(utc, 6, line);
    core_wasm::f64_const(chunk, line, MS_PER_SECOND);
    chunk.emit_op(Op::F64_DIV, line);
}

/// A monotonically-shaped clock reading in seconds. Stack: `[]` → `[num]`.
pub fn emit_clock_seconds(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let now = chunks[current].add_import("ecma:date", "now");
    chunks[current].emit_call(now, 0, line);
    core_wasm::f64_const(&mut chunks[current], line, MS_PER_SECOND);
    chunks[current].emit_op(Op::F64_DIV, line);
}

/// Same reading, in nanoseconds (integer). Stack: `[]` → `[num]`.
pub fn emit_clock_ns(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let now = chunks[current].add_import("ecma:date", "now");
    chunks[current].emit_call(now, 0, line);
    core_wasm::f64_const(&mut chunks[current], line, 1_000_000.0);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_FLOOR, line);
}
