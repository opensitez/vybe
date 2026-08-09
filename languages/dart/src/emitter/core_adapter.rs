//! Dart core library adapters for Duration, DateTime, and Uri.

use std::sync::Arc;
use vybe_compiler::primitives::instructions::{core_wasm, host};
use vybe_compiler::primitives::{collections, reflection, url};
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

fn key(chunk: &mut Chunk, name: &str) -> u16 {
    chunk.add_constant(Value::String(Arc::from(name)))
}

fn set_field(chunk: &mut Chunk, name: &str, line: u32) {
    let k = key(chunk, name);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, k, line);
}

fn get_field(chunk: &mut Chunk, name: &str, line: u32) {
    let k = key(chunk, name);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, k, line);
}

fn emit_slot_is_bigint(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    host::emit(chunk, "wasm:js-bigint", "test", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
}

fn emit_bigint_i32(chunks: &mut [Chunk], current: usize, value: i32, line: u32) {
    chunks[current].emit_i32_const(value, line);
    host::emit(&mut chunks[current], "ecma:bigint", "BigInt", 1, line);
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

fn stamp_runtime_type(
    chunk: &mut Chunk,
    type_name: &str,
    kind: reflection::ReflectKind,
    line: u32,
) {
    set_string(chunk, reflection::FIELD_TYPE, type_name, line);
    set_string(chunk, reflection::FIELD_TYPE_NAME, type_name, line);
    set_string(chunk, reflection::FIELD_KIND, kind.as_str(), line);
}

fn date_get(chunks: &mut [Chunk], current: usize, ms_slot: u16, getter: &'static str, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    host::emit(&mut chunks[current], "ecma:date", getter, 1, line);
}

fn wrap_duration_ms(chunk: &mut Chunk, line: u32) {
    let ms = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, ms, line);
    obj_new(chunk, line);
    stamp_runtime_type(chunk, "Duration", reflection::ReflectKind::Object, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    set_field(chunk, "inMilliseconds", line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    // Spans from `primitives::datetime`; `MS_PER_SECOND` doubles as the
    // millisecond→microsecond factor, Dart's one-tick-finer resolution.
    chunk.emit_f64_const(vybe_compiler::primitives::datetime::MS_PER_SECOND, line);
    chunk.emit_op(Op::F64_MUL, line);
    set_field(chunk, "inMicroseconds", line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    chunk.emit_f64_const(vybe_compiler::primitives::datetime::MS_PER_SECOND, line);
    chunk.emit_op(Op::F64_DIV, line);
    set_field(chunk, "inSeconds", line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    chunk.emit_f64_const(vybe_compiler::primitives::datetime::MS_PER_MINUTE, line);
    chunk.emit_op(Op::F64_DIV, line);
    set_field(chunk, "inMinutes", line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    chunk.emit_f64_const(vybe_compiler::primitives::datetime::MS_PER_HOUR, line);
    chunk.emit_op(Op::F64_DIV, line);
    set_field(chunk, "inHours", line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    chunk.emit_f64_const(vybe_compiler::primitives::datetime::MS_PER_DAY, line);
    chunk.emit_op(Op::F64_DIV, line);
    set_field(chunk, "inDays", line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    chunk.emit_f64_const(0.0, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
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

/// Is the value in `slot` a `type_name`? Stack: [] → [i32].
///
/// Asks IDENTITY first — `emit_is_instance_of` is the shared rtt test, the same
/// one `x is T` compiles to — and only then falls back to the `__type` STRING a
/// legacy wrapper carries. Both, because the two shapes coexist during this
/// migration: `Duration` is a real class now, so it has an rtt and no `__type`,
/// while `DateTime` is still an anonymous struct with a `__type` and no rtt.
///
/// Keeping only the string compare is what made `d.abs()` and
/// `d.compareTo(other)` answer nothing the moment Duration became a class:
/// receiver-blind `[value_methods]` adapters serve every type through this one
/// question, so it has to recognise both.
fn emit_slot_is_type(chunks: &mut [Chunk], current: usize, slot: u16, type_name: &str, line: u32) {
    reflection::emit_is_instance_of(chunks, current, slot, type_name, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    get_field(&mut chunks[current], reflection::FIELD_TYPE, line);
    chunks[current].emit_string_const(type_name, line);
    chunks[current].emit_op(Op::EQ, line);
    chunks[current].emit_op(Op::I32_OR, line);
}

pub fn emit_dart_abs(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    emit_slot_is_bigint(&mut chunks[current], value, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    crate::emitter::string_adapter::emit_dart_bigint_abs(chunks, current, line);
    chunks[current].emit_else(line);
    emit_slot_is_type(chunks, current, value, "Duration", line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    emit_duration_abs(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_op(Op::F64_ABS, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_num_floor(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::F64_FLOOR, line);
}

pub fn emit_num_ceil(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::F64_CEIL, line);
}

pub fn emit_num_round(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:math", "round", 1, line);
}

pub fn emit_num_truncate(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::F64_TRUNC, line);
}

pub fn emit_num_to_double(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
}

pub fn emit_num_remainder(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::math::emit_c_fmod(&mut chunks[current], line);
}

pub fn emit_num_is_negative(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    emit_slot_is_type(chunks, current, value, "Duration", line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    get_field(&mut chunks[current], "isNegative", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_f64_const(0.0, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

pub fn emit_num_is_infinite(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    host::emit(&mut chunks[current], "ecma:number", "isFinite", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    host::emit(&mut chunks[current], "ecma:number", "isNaN", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);

    chunks[current].emit_op(Op::I32_AND, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_num_sign(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(-1.0, line);
    chunks[current].emit_else(line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_duration_negate(chunks: &mut [Chunk], current: usize, line: u32) {
    let slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
    duration_ms_from_obj(&mut chunks[current], slot, line);
    chunks[current].emit_f64_const(-1.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    wrap_duration_ms(&mut chunks[current], line);
}

/// Lower `DateTime(year[, month[, day[, hour[, minute[, second]]]]])` onto
/// `ecma:date.UTC`. Dart §`DateTime` gives every component after `year` a
/// default — `month` and `day` are **1**, the time parts are 0 — so only the
/// `argc` values actually on the stack may be popped; popping a fixed six (or
/// even a fixed three) would consume operands that were never pushed and read
/// whatever happened to be underneath (`DateTime(2000)` → undefined,
/// `DateTime(2000, 1)` → year 2066).
fn utc_from_stack(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = chunks[current].alloc_scratch(6);
    // Slots in ARGUMENT order, so index i is the i-th constructor parameter.
    let slots = [
        base,     // year
        base + 1, // month
        base + 2, // day
        base + 3, // hour
        base + 4, // minute
        base + 5, // second
    ];
    let supplied = (argc as usize).min(slots.len());
    // The last argument is on top, so pop right-to-left.
    for i in (0..supplied).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, slots[i], line);
    }
    // Fill in the omitted components with Dart's defaults.
    for (i, slot) in slots.iter().enumerate().skip(supplied) {
        // month/day are 1-based; hour/minute/second start at 0.
        let default = if i <= 2 { 1.0 } else { 0.0 };
        chunks[current].emit_f64_const(default, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, *slot, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    // `ecma:date.UTC` takes a 0-based month; Dart's is 1-based.
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    for slot in &slots[2..] {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *slot, line);
    }
    host::emit(&mut chunks[current], "ecma:date", "UTC", 6, line);
}

// `wrap_datetime_ms` is GONE. It built an anonymous `ecma:object`, stamped a
// `__type` STRING of "DateTime" on it, and eagerly filled seven calendar parts
// with `ecma:date` getters. `DateTime` is a CLASS now
// (`core_classes/datetime.rs`): the same seven parts are fields filled in its
// constructor, and the type is an rtt rather than a string two other emits had
// to read to find out what they were holding.

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
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    get_field(chunk, "inMilliseconds", line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    host::emit(chunk, "wasm:js-undefined", "test", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_compare_slots(chunk: &mut Chunk, left: u16, right: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, left, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(-1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, left, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(1, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// `__dart_date_month(ms)` — the UTC month in Dart's ONE-BASED numbering.
///
/// `MonthIndexing` is `vybe_ast::datetime` vocabulary, not arithmetic to
/// respell: the host getter is zero-based and the +1 is the convention.
pub fn emit_date_month(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:date", "getUTCMonth", 1, line);
    // The host getter is `MonthIndexing::ZeroBased`; Dart is `OneBased`.
    if DART_MONTHS != vybe_ast::datetime::MonthIndexing::ZeroBased {
        chunks[current].emit_f64_const(1.0, line);
        chunks[current].emit_op(Op::F64_ADD, line);
    }
}

/// Dart numbers months 1–12.
const DART_MONTHS: vybe_ast::datetime::MonthIndexing = vybe_ast::datetime::MonthIndexing::OneBased;

/// `__dart_date_weekday(ms)` — the UTC weekday with Monday = 1.
///
/// Straight through `primitives::datetime::emit_weekday_in_base`, which owns
/// the three bases every language picks from. The emitter this replaces spelled
/// the Sunday-zero fixup as an inline `if dow != 0 { dow } else { 7 }`.
pub fn emit_date_weekday(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:date", "getUTCDay", 1, line);
    vybe_compiler::primitives::datetime::emit_weekday_in_base(
        &mut chunks[current],
        vybe_ast::datetime::WeekdayBase::MondayOne,
        line,
    );
}

// `emit_datetime_new` / `_add` / `_subtract` / `_difference` / `_diff_ms` /
// `_is_before` / `_is_after` / `_same_moment` are GONE — every one of them is a
// member on the `DateTime` CLASS now, spelled in Dart.

/// `[value_methods] add` — the LIST push, and nothing else.
///
/// This used to open by reading `__type` off the receiver and comparing it to
/// the string `"DateTime"`, branching between advancing a moment and appending
/// to a list. That test is what a wrapper forces on every shared entry point
/// that might receive one. `DateTime` declares its own `add` now, so a moment
/// never reaches here.
pub fn emit_dart_add(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_push(chunks, current, line);
}

pub fn emit_compare_to(chunks: &mut [Chunk], current: usize, line: u32) {
    let right = chunks[current].alloc_scratch(5);
    let left = right + 1;
    let r = right + 2;
    let l = right + 3;
    let method = right + 4;
    chunks[current].emit_op_u16(Op::LOCAL_SET, right, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left, line);
    emit_slot_is_bigint(&mut chunks[current], left, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right, line);
    host::emit(&mut chunks[current], "ecma:bigint", "lt", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right, line);
    host::emit(&mut chunks[current], "ecma:bigint", "gt", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    host::emit(&mut chunks[current], "wasm:js-string", "test", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right, line);
    host::emit(&mut chunks[current], "wasm:js-string", "compare", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    host::emit(&mut chunks[current], "wasm:js-number", "test", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_compare_slots(&mut chunks[current], left, right, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    get_field(&mut chunks[current], "compareTo", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, method, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, method, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, method, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 2, 1, line);
    chunks[current].emit_else(line);
    comparable_value_from_obj(&mut chunks[current], left, l, line);
    comparable_value_from_obj(&mut chunks[current], right, r, line);
    emit_compare_slots(&mut chunks[current], l, r, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

// `Uri` is a CLASS (`core_classes/uri.rs`). What used to live here — a
// `wrap_url` that read WHATWG properties and stripped `:` / `?` / `#` by hand
// with `ecma:string.replace` — was `primitives::url`'s `Trim::TrailingColon`
// and `Trim::LeadingChar` written a second time. These readers are the whole
// Dart surface onto that primitive now: one parse, one component model.
//
// The shape is java's (`__j_url_parse` + `__j_url_*` getters, `platforms/jvm`),
// which `emit_component_of` was written for: *"a PROFILE builtin instead
// receives its argument on the stack."*
//
// `ParseMode::Syntactic` with `lowercase_scheme`, not WHATWG. Dart's
// `Uri.parse` is an RFC 3986 SPLIT: it accepts a relative reference
// (`Uri.parse('/a/b').host` is `''`), where WHATWG `new URL` throws, and it
// does not fold host case or drop ports. `emit_parse_syntactic` re-attaches
// the punctuation WHATWG carries, so the same component readers serve both.
fn dart_parse_options() -> url::ParseOptions {
    url::ParseOptions {
        mode: url::ParseMode::Syntactic,
        lowercase_scheme: true }
}

/// `__dart_url_parse(s)` — the ONE parse. Every component read below takes the
/// object this returns.
pub fn emit_url_parse(chunks: &mut [Chunk], current: usize, line: u32) {
    url::emit_parse(chunks, current, dart_parse_options(), line);
}

/// `__dart_url_<field>(parsed)` — one canonical component, trimmed by the
/// shared reader.
pub fn emit_url_component(
    chunks: &mut [Chunk],
    current: usize,
    field: url::UrlField,
    line: u32,
) {
    url::emit_component_of(chunks, current, field, line);
}

/// `__dart_url_decode(s)` — RFC 3986 percent-decoding.
///
/// The walker fold this replaces spelled it `text.replace("%20", " ")`, which
/// decoded exactly one sequence. Every other escape came through verbatim.
pub fn emit_url_decode(chunks: &mut [Chunk], current: usize, line: u32) {
    url::emit_percent_decode(chunks, current, url::PercentOptions::rfc3986(), line);
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
