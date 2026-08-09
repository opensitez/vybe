//! Behaviour tests for `ecma:date` host imports.
//!
//! Reference: ECMA-262 §21.4 Date.
//!
//! Each test covers a distinct behaviour.

use std::sync::Arc;
use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};

static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn push_arg(vm: &mut VM, chunk: &mut Chunk, value: Value) {
    match value {
        Value::I32(n) => chunk.emit_i32_const(n, 0),
        Value::I64(n) => chunk.emit_i64_const(n, 0),
        Value::F32(f) => chunk.emit_f32_const(f, 0),
        Value::F64(f) => chunk.emit_f64_const(f, 0),
        Value::Bool(b) => chunk.emit_bool_const(b, 0),
        Value::String(s) => chunk.emit_string_const(&s, 0),
        Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0),
        other => {
            let global = format!(
                "__test_arg_{}",
                TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
            );
            vm.globals.insert(global.clone(), other);
            let ci = chunk.intern_string_constant(&global);
            chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
        }
    }
}

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let mut chunk = Chunk::new("<ecma-date-test>");
    let import_idx = chunk.add_import("ecma:date", name);
    let argc = args.len() as u8;
    for value in args {
        push_arg(&mut vm, &mut chunk, value);
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("VM run failed")
}

fn invoke_value(name: &str, args: Vec<Value>) -> Value {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let mut chunk = Chunk::new("<ecma-value-date-test>");
    let import_idx = chunk.add_import("ecma:value", name);
    let argc = args.len() as u8;
    for value in args {
        push_arg(&mut vm, &mut chunk, value);
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("VM run failed")
}

fn as_num(v: &Value) -> f64 {
    match v {
        Value::I32(n) => *n as f64,
        Value::I64(n) => *n as f64,
        Value::F64(f) => *f,
        _ => panic!("expected numeric, got {:?}", v),
    }
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn epoch_date() -> Value {
    invoke("new", vec![Value::F64(0.0)])
}

// ── Date.now ──────────────────────────────────────────────────────────────────

#[test]
fn now_returns_milliseconds_past_unix_epoch_in_year_2000_plus() {
    // Jan 1 2001 = 978307200000 ms. Any modern clock must exceed this.
    let ms = as_num(&invoke("now", vec![]));
    assert!(
        ms > 978_307_200_000.0,
        "Date.now() must exceed year-2001 timestamp"
    );
}

#[test]
fn now_does_not_decrease_between_two_calls() {
    let t1 = as_num(&invoke("now", vec![]));
    let t2 = as_num(&invoke("now", vec![]));
    assert!(t2 >= t1);
}

// ── Date constructor — from milliseconds ──────────────────────────────────────

#[test]
fn new_from_zero_ms_is_unix_epoch() {
    let d = epoch_date();
    assert_eq!(as_num(&invoke("getFullYear", vec![d])) as i32, 1970);
}

#[test]
fn get_time_round_trips_constructor_milliseconds() {
    let ms = 1_700_000_000_000.0_f64;
    let d = invoke("new", vec![Value::F64(ms)]);
    assert_eq!(as_num(&invoke("getTime", vec![d])), ms);
}

#[test]
fn value_of_equals_get_time() {
    let ms = 1_000_000_000_000.0_f64;
    let d = invoke("new", vec![Value::F64(ms)]);
    assert_eq!(
        as_num(&invoke("valueOf", vec![d.clone()])),
        as_num(&invoke("getTime", vec![d]))
    );
}

// ── Date constructor — component-based ───────────────────────────────────────

#[test]
fn month_is_zero_indexed_january_is_zero() {
    // new Date(2000, 0, 1) = Jan 1 2000; month 0 = January.
    let d = invoke("new", vec![Value::I32(2000), Value::I32(0), Value::I32(1)]);
    assert_eq!(as_num(&invoke("getMonth", vec![d])) as i32, 0);
}

#[test]
fn get_date_returns_day_of_month_one_indexed() {
    let d = invoke("new", vec![Value::I32(2024), Value::I32(5), Value::I32(15)]);
    assert_eq!(as_num(&invoke("getDate", vec![d])) as i32, 15);
}

#[test]
fn get_day_returns_zero_for_sunday() {
    // Jan 1 1978 was a Sunday.
    let d = invoke("new", vec![Value::I32(1978), Value::I32(0), Value::I32(1)]);
    assert_eq!(as_num(&invoke("getDay", vec![d])) as i32, 0);
}

// ── toISOString ───────────────────────────────────────────────────────────────

#[test]
fn to_iso_string_of_epoch_is_canonical_utc_format() {
    let d = epoch_date();
    match invoke("toISOString", vec![d]) {
        Value::String(s) => assert_eq!(s.as_ref(), "1970-01-01T00:00:00.000Z"),
        other => panic!("expected string, got {:?}", other),
    }
}

#[test]
fn to_iso_string_always_ends_with_z_for_utc() {
    let d = invoke("new", vec![Value::F64(1_700_000_000_000.0)]);
    match invoke("toISOString", vec![d]) {
        Value::String(s) => assert!(s.ends_with('Z'), "must end with Z: {s}"),
        other => panic!("expected string, got {:?}", other),
    }
}

#[test]
fn to_json_returns_same_as_to_iso_string() {
    let d = epoch_date();
    let iso = invoke("toISOString", vec![d.clone()]);
    let json = invoke("toJSON", vec![d]);
    assert_eq!(iso, json);
}

// ── String representations ────────────────────────────────────────────────────

#[test]
fn to_string_returns_a_non_empty_string() {
    let d = epoch_date();
    match invoke("toString", vec![d]) {
        Value::String(s) => assert!(!s.is_empty()),
        other => panic!("expected string, got {:?}", other),
    }
}

#[test]
fn to_date_string_does_not_include_time_component() {
    // toDateString like "Thu Jan 01 1970" contains no colon (no time).
    let d = epoch_date();
    match invoke("toDateString", vec![d]) {
        Value::String(s) => assert!(!s.contains(':'), "toDateString must not include time: {s}"),
        other => panic!("expected string, got {:?}", other),
    }
}

#[test]
fn value_invoke_method_locale_date_string_on_date_object() {
    let d = epoch_date();
    match invoke_value("invokeMethod", vec![d, s("toLocaleDateString")]) {
        Value::String(text) => assert!(!text.is_empty()),
        other => panic!("expected string, got {:?}", other),
    }
}

#[test]
fn to_utc_string_contains_gmt_or_utc_marker() {
    let d = epoch_date();
    match invoke("toUTCString", vec![d]) {
        Value::String(s) => assert!(
            s.contains("GMT") || s.contains("UTC"),
            "toUTCString must reference UTC: {s}"
        ),
        other => panic!("expected string, got {:?}", other),
    }
}

// ── Date.UTC static ───────────────────────────────────────────────────────────

#[test]
fn utc_1970_jan_1_returns_zero() {
    let result = invoke("UTC", vec![Value::I32(1970), Value::I32(0), Value::I32(1)]);
    assert_eq!(as_num(&result), 0.0);
}

#[test]
fn utc_returns_millisecond_timestamp_for_any_date() {
    let result = invoke("UTC", vec![Value::I32(2000), Value::I32(0), Value::I32(1)]);
    // Jan 1 2000 00:00:00 UTC = 946684800000 ms
    assert!((as_num(&result) - 946_684_800_000.0).abs() < 1000.0);
}

// ── Date.parse ────────────────────────────────────────────────────────────────

#[test]
fn parse_iso_string_returns_its_millisecond_timestamp() {
    let result = invoke(
        "parse",
        vec![Value::String(Arc::from("1970-01-01T00:00:00.000Z"))],
    );
    assert_eq!(as_num(&result), 0.0);
}

// ── Setters round-trip with getters ──────────────────────────────────────────

#[test]
fn set_full_year_is_reflected_by_get_full_year() {
    let d = epoch_date();
    invoke("setFullYear", vec![d.clone(), Value::I32(2042)]);
    assert_eq!(as_num(&invoke("getFullYear", vec![d])) as i32, 2042);
}

#[test]
fn set_hours_does_not_change_date_component() {
    let d = invoke("new", vec![Value::I32(2000), Value::I32(0), Value::I32(15)]);
    invoke("setHours", vec![d.clone(), Value::I32(12)]);
    assert_eq!(as_num(&invoke("getDate", vec![d])) as i32, 15);
}

// ── getTimezoneOffset ─────────────────────────────────────────────────────────

#[test]
fn get_timezone_offset_is_integer_minutes_within_plausible_range() {
    let d = epoch_date();
    let offset = as_num(&invoke("getTimezoneOffset", vec![d]));
    // Valid UTC offsets: -720 (UTC-12) to +840 (UTC+14), in minutes.
    assert!(
        offset >= -840.0 && offset <= 720.0,
        "offset out of range: {offset}"
    );
    assert_eq!(offset, offset.trunc(), "offset must be integer minutes");
}

// ── fromUnixSeconds / toUnixSeconds ──────────────────────────────────────────

#[test]
fn from_unix_seconds_zero_is_epoch() {
    let d = invoke("fromUnixSeconds", vec![Value::F64(0.0)]);
    assert_eq!(as_num(&invoke("getFullYear", vec![d])) as i32, 1970);
}

#[test]
fn to_unix_seconds_round_trips_from_unix_seconds() {
    let secs = 1_000_000.0_f64;
    let d = invoke("fromUnixSeconds", vec![Value::F64(secs)]);
    let back = as_num(&invoke("toUnixSeconds", vec![d]));
    assert!((back - secs).abs() < 1.0);
}

// ── getUTC* methods return UTC components ─────────────────────────────────────

#[test]
fn get_utc_full_year_of_epoch_is_1970() {
    // ECMA-262 §21.4.4.14: getUTCFullYear always uses UTC, never local time.
    let d = invoke("new", vec![Value::F64(0.0)]);
    assert_eq!(as_num(&invoke("getUTCFullYear", vec![d])) as i32, 1970);
}

#[test]
fn get_utc_month_of_epoch_is_january() {
    let d = invoke("new", vec![Value::F64(0.0)]);
    assert_eq!(as_num(&invoke("getUTCMonth", vec![d])) as i32, 0);
}

#[test]
fn get_utc_hours_of_epoch_is_zero() {
    // Unix epoch 0 = midnight UTC; getUTCHours must return 0.
    let d = invoke("new", vec![Value::F64(0.0)]);
    assert_eq!(as_num(&invoke("getUTCHours", vec![d])) as i32, 0);
}

#[test]
fn get_hours_and_get_minutes_return_numeric_values_in_range() {
    // §21.4.4.11/4.12: getHours ∈ [0,23]; getMinutes ∈ [0,59].
    let d = invoke("now", vec![]);
    let d_obj = invoke("new", vec![d]);
    let h = as_num(&invoke("getHours", vec![d_obj.clone()])) as i32;
    let m = as_num(&invoke("getMinutes", vec![d_obj])) as i32;
    assert!((0..=23).contains(&h));
    assert!((0..=59).contains(&m));
}

#[test]
fn get_milliseconds_is_in_range_zero_to_999() {
    // §21.4.4.13: getMilliseconds ∈ [0,999].
    let d = invoke("new", vec![Value::F64(1_700_000_000_123.0)]);
    let ms = as_num(&invoke("getMilliseconds", vec![d])) as i32;
    assert!((0..=999).contains(&ms));
}

// ── Invalid Date ──────────────────────────────────────────────────────────────

#[test]
fn invalid_date_to_iso_string_is_nan_string() {
    // ECMA-262 §21.4.4.36: If the Date is invalid, toISOString throws or returns "Invalid Date".
    let d = invoke("new", vec![Value::F64(f64::NAN)]);
    let result = invoke("toISOString", vec![d]);
    // Either throws (Undefined returned on error) or returns "Invalid Date".
    match result {
        Value::String(s) => assert!(s.contains("Invalid") || s.contains("NaN")),
        Value::Undefined => {}
        other => panic!("unexpected: {:?}", other),
    }
}

// ── Missing Date setters ──────────────────────────────────────────────────────

#[test]
fn set_month_changes_month_component() {
    // ECMA-262 §21.4.4.27: setMonth(m) replaces the month (0-indexed).
    let d = invoke("new", vec![Value::I32(2024), Value::I32(0), Value::I32(15)]);
    invoke("setMonth", vec![d.clone(), Value::I32(5)]);
    assert_eq!(as_num(&invoke("getMonth", vec![d])) as i32, 5);
}

#[test]
fn set_date_changes_day_of_month() {
    // ECMA-262 §21.4.4.21: setDate(d) sets the day of the month.
    let d = invoke("new", vec![Value::I32(2024), Value::I32(0), Value::I32(1)]);
    invoke("setDate", vec![d.clone(), Value::I32(20)]);
    assert_eq!(as_num(&invoke("getDate", vec![d])) as i32, 20);
}

#[test]
fn set_minutes_changes_minute_component() {
    // ECMA-262 §21.4.4.26: setMinutes(m) sets the minutes.
    let d = invoke("new", vec![Value::F64(0.0)]);
    invoke("setMinutes", vec![d.clone(), Value::I32(30)]);
    assert_eq!(as_num(&invoke("getMinutes", vec![d])) as i32, 30);
}

#[test]
fn set_seconds_changes_second_component() {
    let d = invoke("new", vec![Value::F64(0.0)]);
    invoke("setSeconds", vec![d.clone(), Value::I32(45)]);
    assert_eq!(as_num(&invoke("getSeconds", vec![d])) as i32, 45);
}

#[test]
fn set_milliseconds_changes_millisecond_component() {
    let d = invoke("new", vec![Value::F64(0.0)]);
    invoke("setMilliseconds", vec![d.clone(), Value::I32(500)]);
    assert_eq!(as_num(&invoke("getMilliseconds", vec![d])) as i32, 500);
}

#[test]
fn set_time_resets_to_given_timestamp() {
    // ECMA-262 §21.4.4.32: setTime(t) sets internal [[DateValue]] to t ms.
    let d = invoke("new", vec![Value::F64(0.0)]);
    let new_time = 1_700_000_000_000.0_f64;
    invoke("setTime", vec![d.clone(), Value::F64(new_time)]);
    assert_eq!(as_num(&invoke("getTime", vec![d])), new_time);
}

// ── toTimeString / toLocaleString ─────────────────────────────────────────────

#[test]
fn to_time_string_contains_colon_separators() {
    // ECMA-262 §21.4.4.39: toTimeString formats the time portion (has colons).
    let d = epoch_date();
    match invoke("toTimeString", vec![d]) {
        Value::String(s) => assert!(s.contains(':'), "toTimeString must contain ':': {s}"),
        other => panic!("expected string, got {:?}", other),
    }
}

#[test]
fn to_locale_string_returns_non_empty_string() {
    // ECMA-262 §21.4.4.38: toLocaleString returns a locale-sensitive string.
    let d = invoke("new", vec![Value::F64(1_700_000_000_000.0)]);
    match invoke("toLocaleString", vec![d]) {
        Value::String(s) => assert!(!s.is_empty()),
        other => panic!("expected string, got {:?}", other),
    }
}

// ── toLocaleDateString / toLocaleTimeString ───────────────────────────────────

#[test]
fn to_locale_date_string_returns_non_empty_string() {
    // ECMA-262 §21.4.4.36: toLocaleDateString returns locale-formatted date portion.
    let d = invoke("new", vec![Value::F64(1_700_000_000_000.0)]);
    match invoke("toLocaleDateString", vec![d]) {
        Value::String(s) => assert!(!s.is_empty()),
        other => panic!("expected string, got {:?}", other),
    }
}

#[test]
fn to_locale_time_string_returns_non_empty_string() {
    // ECMA-262 §21.4.4.37: toLocaleTimeString returns locale-formatted time portion.
    let d = invoke("new", vec![Value::F64(1_700_000_000_000.0)]);
    match invoke("toLocaleTimeString", vec![d]) {
        Value::String(s) => assert!(!s.is_empty()),
        other => panic!("expected string, got {:?}", other),
    }
}

// ── UTC getter/setter methods (§21.4.4.x) ────────────────────────────────────

#[test]
fn get_utc_date_returns_day_of_month() {
    // 2024-03-15 = year 2024, month 2 (0-indexed), day 15
    let d = invoke("new", vec![Value::I32(2024), Value::I32(2), Value::I32(15)]);
    assert_eq!(invoke("getUTCDate", vec![d]).as_i32(), 15);
}

#[test]
fn get_utc_day_returns_valid_weekday() {
    let d = invoke("new", vec![Value::I32(2024), Value::I32(2), Value::I32(15)]);
    let day = invoke("getUTCDay", vec![d]).as_i32();
    assert!((0..7).contains(&day));
}

#[test]
fn get_utc_milliseconds_returns_ms_component() {
    // Construct date with known ms offset via timestamp
    let ms = invoke(
        "parse",
        vec![Value::String(Arc::from("2024-01-01T00:00:00.250Z"))],
    );
    let d = invoke("new", vec![ms]);
    assert_eq!(invoke("getUTCMilliseconds", vec![d]).as_i32(), 250);
}

#[test]
fn get_utc_minutes_returns_minute_component() {
    let ms = invoke(
        "parse",
        vec![Value::String(Arc::from("2024-01-01T12:34:00Z"))],
    );
    let d = invoke("new", vec![ms]);
    assert_eq!(invoke("getUTCMinutes", vec![d]).as_i32(), 34);
}

#[test]
fn get_utc_seconds_returns_second_component() {
    let ms = invoke(
        "parse",
        vec![Value::String(Arc::from("2024-01-01T00:00:42Z"))],
    );
    let d = invoke("new", vec![ms]);
    assert_eq!(invoke("getUTCSeconds", vec![d]).as_i32(), 42);
}

#[test]
fn set_utc_milliseconds_updates_ms_field() {
    let d = invoke("new", vec![Value::F64(0.0)]); // epoch, 0ms
    invoke("setUTCMilliseconds", vec![d.clone(), Value::I32(750)]);
    assert_eq!(invoke("getUTCMilliseconds", vec![d]).as_i32(), 750);
}
