//! Behaviour tests for `wasi:clocks` proposal-shaped host imports.

use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};

use vybe_bytecode::value::Value;
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn invoke(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-clocks-test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    for value in args {
        let constant = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, constant, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn invoke_result(module: &str, name: &str, args: Vec<Value>) -> Result<Value, String> {
    let mut chunk = Chunk::new("<wasi-clocks-test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    for value in args {
        let constant = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, constant, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).map_err(|error| error.message)
}

fn has_import(module: &str, name: &str) -> bool {
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(module.to_string(), name.to_string()))
}

fn prop(value: &Value, key: &str) -> Value {
    if let Value::Object(object) = value {
        let object = object.lock().unwrap();
        if let Some(value) = object.properties.get(key) {
            return value.clone();
        }
    }
    Value::Null
}

fn f64_prop(value: &Value, key: &str) -> f64 {
    prop(value, key).as_f64()
}

fn s(text: &str) -> Value {
    Value::String(std::sync::Arc::from(text))
}

#[test]
fn wall_clock_now_returns_current_datetime_record() {
    let before = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap()
        .as_secs_f64();
    let now = invoke("wasi:clocks/wall-clock", "now", vec![]);
    let after = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap()
        .as_secs_f64();

    let seconds = f64_prop(&now, "seconds");
    let nanos = f64_prop(&now, "nanoseconds");
    assert!(
        seconds >= before.floor(),
        "seconds should be current epoch seconds"
    );
    assert!(
        seconds <= after.ceil(),
        "seconds should not be in the future"
    );
    assert!(
        (0.0..1_000_000_000.0).contains(&nanos),
        "nanoseconds should be a u32 tick"
    );
}

#[test]
fn wall_clock_resolution_returns_single_nanosecond_tick() {
    let resolution = invoke("wasi:clocks/wall-clock", "resolution", vec![]);
    assert_eq!(prop(&resolution, "seconds"), Value::F64(0.0));
    assert_eq!(prop(&resolution, "nanoseconds"), Value::F64(1.0));
}

#[test]
fn monotonic_clock_now_is_non_decreasing() {
    let first = invoke("wasi:clocks/monotonic-clock", "now", vec![]).as_f64();
    std::thread::sleep(Duration::from_millis(5));
    let second = invoke("wasi:clocks/monotonic-clock", "now", vec![]).as_f64();
    let resolution = invoke("wasi:clocks/monotonic-clock", "resolution", vec![]).as_f64();

    assert!(second >= first, "monotonic time should not go backwards");
    assert!(resolution >= 1.0, "resolution should be positive");
}

#[test]
fn legacy_now_returns_milliseconds_near_system_clock() {
    let before = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap()
        .as_millis() as f64;
    let now = invoke("wasi:clocks", "now", vec![]).as_f64();
    let after = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap()
        .as_millis() as f64;

    assert!(
        now >= before,
        "legacy now should not predate the call window"
    );
    assert!(now <= after, "legacy now should be within the call window");
}

#[test]
fn legacy_sleep_blocks_for_requested_duration() {
    let start = Instant::now();
    let result = invoke("wasi:clocks", "sleep", vec![Value::F64(20.0)]);
    let elapsed = start.elapsed();

    assert!(matches!(result, Value::Null));
    assert!(
        elapsed >= Duration::from_millis(15),
        "sleep should block for roughly the requested duration"
    );
}

#[test]
fn legacy_to_iso_string_formats_unix_epoch() {
    let result = invoke("wasi:clocks", "toISOString", vec![Value::F64(0.0)]);
    assert_eq!(result, s("1970-01-01T00:00:00.000Z"));
}

#[test]
fn monotonic_clock_resolution_is_exactly_one_nanosecond() {
    let resolution = invoke("wasi:clocks/monotonic-clock", "resolution", vec![]);
    assert_eq!(resolution, Value::F64(1.0));
}

#[test]
fn legacy_hrtime_is_non_decreasing() {
    let first = invoke("wasi:clocks", "hrtime", vec![]).as_f64();
    std::thread::sleep(Duration::from_millis(2));
    let second = invoke("wasi:clocks", "hrtime", vec![]).as_f64();
    assert!(second >= first);
}

#[test]
fn stopwatch_new_starts_stopped_with_zero_elapsed() {
    let stopwatch = invoke("wasi:clocks", "stopwatchNew", vec![]);
    assert_eq!(prop(&stopwatch, "isrunning"), Value::Bool(false));
    let elapsed = invoke("wasi:clocks", "stopwatchElapsed", vec![stopwatch]);
    assert_eq!(elapsed, Value::F64(0.0));
}

#[test]
fn stopwatch_start_then_elapsed_becomes_positive() {
    let stopwatch = invoke("wasi:clocks", "stopwatchNew", vec![]);
    assert!(matches!(
        invoke("wasi:clocks", "stopwatchStart", vec![stopwatch.clone()]),
        Value::Null
    ));
    std::thread::sleep(Duration::from_millis(10));
    let elapsed = invoke("wasi:clocks", "stopwatchElapsed", vec![stopwatch.clone()]).as_f64();
    assert!(elapsed >= 5.0);
    assert_eq!(prop(&stopwatch, "isrunning"), Value::Bool(true));
}

#[test]
fn stopwatch_stop_freezes_elapsed_value() {
    let stopwatch = invoke("wasi:clocks", "stopwatchNew", vec![]);
    let _ = invoke("wasi:clocks", "stopwatchStart", vec![stopwatch.clone()]);
    std::thread::sleep(Duration::from_millis(10));
    let _ = invoke("wasi:clocks", "stopwatchStop", vec![stopwatch.clone()]);
    let first = invoke("wasi:clocks", "stopwatchElapsed", vec![stopwatch.clone()]).as_f64();
    std::thread::sleep(Duration::from_millis(10));
    let second = invoke("wasi:clocks", "stopwatchElapsed", vec![stopwatch.clone()]).as_f64();
    assert_eq!(prop(&stopwatch, "isrunning"), Value::Bool(false));
    assert!(
        (second - first).abs() < 5.0,
        "elapsed should stop increasing once stopped"
    );
}

#[test]
fn stopwatch_reset_clears_accumulated_elapsed() {
    let stopwatch = invoke("wasi:clocks", "stopwatchNew", vec![]);
    let _ = invoke("wasi:clocks", "stopwatchStart", vec![stopwatch.clone()]);
    std::thread::sleep(Duration::from_millis(8));
    let _ = invoke("wasi:clocks", "stopwatchStop", vec![stopwatch.clone()]);
    let _ = invoke("wasi:clocks", "stopwatchReset", vec![stopwatch.clone()]);
    assert_eq!(prop(&stopwatch, "isrunning"), Value::Bool(false));
    assert_eq!(
        invoke("wasi:clocks", "stopwatchElapsed", vec![stopwatch]),
        Value::F64(0.0)
    );
}

#[test]
fn stopwatch_start_while_running_preserves_original_start_time() {
    let stopwatch = invoke("wasi:clocks", "stopwatchNew", vec![]);
    let _ = invoke("wasi:clocks", "stopwatchStart", vec![stopwatch.clone()]);
    std::thread::sleep(Duration::from_millis(6));
    let _ = invoke("wasi:clocks", "stopwatchStart", vec![stopwatch.clone()]);
    std::thread::sleep(Duration::from_millis(6));
    let elapsed = invoke("wasi:clocks", "stopwatchElapsed", vec![stopwatch]).as_f64();
    assert!(elapsed >= 8.0, "second start should not reset elapsed time");
}

#[test]
fn to_iso_string_formats_known_leap_day_timestamp() {
    let result = invoke(
        "wasi:clocks",
        "toISOString",
        vec![Value::F64(951_827_696_789.0)],
    );
    assert_eq!(result, s("2000-02-29T12:34:56.789Z"));
}

#[test]
fn vb_date_extractors_read_epoch_timestamp_components() {
    let timestamp = Value::F64(0.0);
    assert_eq!(
        invoke("wasi:clocks", "vbYear", vec![timestamp.clone()]),
        Value::F64(1970.0)
    );
    assert_eq!(
        invoke("wasi:clocks", "vbMonth", vec![timestamp.clone()]),
        Value::F64(1.0)
    );
    assert_eq!(
        invoke("wasi:clocks", "vbDay", vec![timestamp]),
        Value::F64(1.0)
    );
}

#[test]
fn vb_time_extractors_read_numeric_timestamp_components() {
    let timestamp = Value::F64(3_661.0);
    assert_eq!(
        invoke("wasi:clocks", "vbHour", vec![timestamp.clone()]),
        Value::F64(1.0)
    );
    assert_eq!(
        invoke("wasi:clocks", "vbMinute", vec![timestamp.clone()]),
        Value::F64(1.0)
    );
    assert_eq!(
        invoke("wasi:clocks", "vbSecond", vec![timestamp]),
        Value::F64(1.0)
    );
}

#[test]
fn vb_now_date_and_time_return_string_shapes() {
    let now = invoke("wasi:clocks", "vbNow", vec![]);
    let date = invoke("wasi:clocks", "vbDate", vec![]);
    let time = invoke("wasi:clocks", "vbTime", vec![]);

    let Value::String(now_text) = now else {
        panic!("vbNow should return string")
    };
    let Value::String(date_text) = date else {
        panic!("vbDate should return string")
    };
    let Value::String(time_text) = time else {
        panic!("vbTime should return string")
    };

    assert_eq!(now_text.len(), 19);
    assert_eq!(&now_text[4..5], "-");
    assert_eq!(&date_text[2..3], "/");
    assert_eq!(&date_text[5..6], "/");
    assert_eq!(&time_text[2..3], ":");
    assert_eq!(&time_text[5..6], ":");
}

#[test]
fn vb_timer_returns_seconds_since_midnight_range() {
    let seconds = invoke("wasi:clocks", "vbTimer", vec![]).as_f64();
    assert!((0.0..86_400.0).contains(&seconds));
}

macro_rules! iso_fixture_test {
    ($name:ident, $millis:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let result = invoke("wasi:clocks", "toISOString", vec![Value::F64($millis)]);
            assert_eq!(result, s($expected));
        }
    };
}

macro_rules! vb_component_test {
    ($name:ident, $method:expr, $seconds:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let result = invoke("wasi:clocks", $method, vec![Value::F64($seconds)]);
            assert_eq!(result, Value::F64($expected));
        }
    };
}

iso_fixture_test!(
    to_iso_string_formats_one_hour_one_minute_one_second,
    3_661_000.0,
    "1970-01-01T01:01:01.000Z"
);
iso_fixture_test!(
    to_iso_string_formats_y2k_leap_day_without_millis,
    951_827_696_000.0,
    "2000-02-29T12:34:56.000Z"
);
iso_fixture_test!(
    to_iso_string_formats_one_billion_second_timestamp,
    1_000_000_000_000.0,
    "2001-09-09T01:46:40.000Z"
);
iso_fixture_test!(
    to_iso_string_formats_recent_leap_day_without_millis,
    1_709_210_096_000.0,
    "2024-02-29T12:34:56.000Z"
);

vb_component_test!(
    vb_year_extracts_2000_from_leap_day_fixture,
    "vbYear",
    951_827_696.0,
    2000.0
);
vb_component_test!(
    vb_month_extracts_february_from_leap_day_fixture,
    "vbMonth",
    951_827_696.0,
    2.0
);
vb_component_test!(
    vb_day_extracts_twenty_nine_from_leap_day_fixture,
    "vbDay",
    951_827_696.0,
    29.0
);
vb_component_test!(
    vb_hour_extracts_twelve_from_leap_day_fixture,
    "vbHour",
    951_827_696.0,
    12.0
);
vb_component_test!(
    vb_minute_extracts_thirty_four_from_leap_day_fixture,
    "vbMinute",
    951_827_696.0,
    34.0
);
vb_component_test!(
    vb_second_extracts_fifty_six_from_leap_day_fixture,
    "vbSecond",
    951_827_696.0,
    56.0
);

vb_component_test!(
    vb_year_extracts_2024_from_recent_leap_day_fixture,
    "vbYear",
    1_709_210_096.0,
    2024.0
);
vb_component_test!(
    vb_month_extracts_february_from_recent_leap_day_fixture,
    "vbMonth",
    1_709_210_096.0,
    2.0
);
vb_component_test!(
    vb_day_extracts_twenty_nine_from_recent_leap_day_fixture,
    "vbDay",
    1_709_210_096.0,
    29.0
);
vb_component_test!(
    vb_hour_extracts_twelve_from_recent_leap_day_fixture,
    "vbHour",
    1_709_210_096.0,
    12.0
);
vb_component_test!(
    vb_minute_extracts_thirty_four_from_recent_leap_day_fixture,
    "vbMinute",
    1_709_210_096.0,
    34.0
);
vb_component_test!(
    vb_second_extracts_fifty_six_from_recent_leap_day_fixture,
    "vbSecond",
    1_709_210_096.0,
    56.0
);

#[test]
fn proposal_system_clock_now_import_resolves() {
    assert!(
        invoke_result("wasi:clocks/system-clock", "now", vec![]).is_ok(),
        "wasi:clocks/system-clock.now should be covered by the clocks category"
    );
}

#[test]
fn proposal_system_clock_get_resolution_import_resolves() {
    assert!(
        invoke_result("wasi:clocks/system-clock", "get-resolution", vec![]).is_ok(),
        "wasi:clocks/system-clock.get-resolution should be covered by the clocks category"
    );
}

#[test]
fn proposal_monotonic_clock_get_resolution_import_resolves() {
    assert!(
        invoke_result("wasi:clocks/monotonic-clock", "get-resolution", vec![]).is_ok(),
        "wasi:clocks/monotonic-clock.get-resolution should be covered by the clocks category"
    );
}

#[test]
fn proposal_monotonic_clock_wait_until_import_is_registered() {
    assert!(
        has_import("wasi:clocks/monotonic-clock", "wait-until"),
        "wasi:clocks/monotonic-clock.wait-until should be covered by the clocks category"
    );
}

#[test]
fn proposal_monotonic_clock_wait_for_import_is_registered() {
    assert!(
        has_import("wasi:clocks/monotonic-clock", "wait-for"),
        "wasi:clocks/monotonic-clock.wait-for should be covered by the clocks category"
    );
}
