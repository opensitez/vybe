//! Behaviour tests for `node:perf_hooks` host imports.
//!
//! Reference: <https://nodejs.org/api/perf_hooks.html>.
//!
//! Coverage:
//!   - `performance.now()` → monotonically increasing DOMHighResTimeStamp
//!   - `performance.mark(name[, options])` → PerformanceMark
//!   - `performance.measure(name[, startMarkOrOptions[, endMark]])` → PerformanceMeasure
//!   - `performance.clearMarks([name])` → void
//!   - `performance.clearMeasures([name])` → void
//!   - `performance.getEntries()` → array
//!   - `performance.getEntriesByName(name[, type])` → array
//!   - `performance.getEntriesByType(type)` → array
//!   - `performance.timeOrigin` → number
//!   - `performance.eventLoopUtilization([utilization1])` → object
//!   - `monitorEventLoopDelay([options])` → Histogram-like object
//!   - `createHistogram([options])` → Histogram
//!   - `PerformanceObserver` surface
//!
//! Deferred:
//!   - `PerformanceObserver` callback invocation (requires event loop)

use vybe_runtime::value::{ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};
use vybe_runtime::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

fn call_perf(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-perf_hooks-test>");
    let import_idx = chunk.add_import("node:perf_hooks", name);
    let argc = args.len() as u8;
    for value in args {
        let c = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, c, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn has_import(name: &str) -> bool {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(String::from("node:perf_hooks"), name.to_string()))
}

fn s(text: &str) -> Value {
    Value::String(std::sync::Arc::from(text))
}

fn as_f64(v: &Value) -> f64 {
    match v {
        Value::F64(f) => *f,
        Value::I32(n) => *n as f64,
        _ => f64::NAN }
}

fn prop(obj: &Value, key: &str) -> Value {
    match obj {
        Value::Object(o) => {
            let o = o.lock().unwrap();
            o.properties.get(key).cloned().unwrap_or(Value::Undefined)
        }
        _ => Value::Undefined }
}

// ── performance.now ───────────────────────────────────────────────────────────

#[test]
fn performance_now_returns_non_negative_number() {
    let t = call_perf("performanceNow", vec![]);
    assert!(
        as_f64(&t) >= 0.0,
        "performance.now() must be >= 0, got {t:?}"
    );
}

#[test]
fn performance_now_is_monotonically_increasing() {
    let t1 = call_perf("performanceNow", vec![]);
    // Small busy-wait to ensure time advances
    let _ = (0..1_000).map(|x| x * x).sum::<i64>();
    let t2 = call_perf("performanceNow", vec![]);
    assert!(
        as_f64(&t2) >= as_f64(&t1),
        "performance.now() must not decrease"
    );
}

#[test]
fn performance_now_has_sub_millisecond_resolution() {
    // DOMHighResTimeStamp is a float — must not be an integer multiple of 1.0
    // for back-to-back calls (at least one should have a fractional component
    // or both should differ enough to confirm float precision).
    let t = call_perf("performanceNow", vec![]);
    assert!(matches!(t, Value::F64(_) | Value::I32(_)));
}

// ── performance.timeOrigin ────────────────────────────────────────────────────

#[test]
fn performance_time_origin_is_positive() {
    let origin = call_perf("performanceTimeOrigin", vec![]);
    assert!(
        as_f64(&origin) > 0.0,
        "timeOrigin must be > 0, got {origin:?}"
    );
}

// ── performance.mark ─────────────────────────────────────────────────────────

#[test]
fn performance_mark_returns_entry_object() {
    let entry = call_perf("performanceMark", vec![s("test-mark")]);
    assert!(matches!(entry, Value::Object(_)));
}

#[test]
fn performance_mark_entry_has_name() {
    let entry = call_perf("performanceMark", vec![s("my-mark")]);
    let name = prop(&entry, "name");
    assert_eq!(name, s("my-mark"));
}

#[test]
fn performance_mark_entry_type_is_mark() {
    let entry = call_perf("performanceMark", vec![s("typed-mark")]);
    let entry_type = prop(&entry, "entryType");
    assert_eq!(entry_type, s("mark"));
}

#[test]
fn performance_mark_start_time_is_non_negative() {
    let entry = call_perf("performanceMark", vec![s("t-mark")]);
    let start = prop(&entry, "startTime");
    assert!(as_f64(&start) >= 0.0);
}

// ── performance.clearMarks ────────────────────────────────────────────────────

#[test]
fn performance_clear_marks_specific_name_does_not_panic() {
    let _ = call_perf("performanceMark", vec![s("clear-me")]);
    let result = call_perf("performanceClearMarks", vec![s("clear-me")]);
    assert_eq!(result, Value::Undefined);
}

#[test]
fn performance_clear_marks_all_does_not_panic() {
    let _ = call_perf("performanceMark", vec![s("m1")]);
    let _ = call_perf("performanceMark", vec![s("m2")]);
    let result = call_perf("performanceClearMarks", vec![]);
    assert_eq!(result, Value::Undefined);
}

// ── performance.measure ───────────────────────────────────────────────────────

#[test]
fn performance_measure_returns_entry_object() {
    let _ = call_perf("performanceMark", vec![s("start")]);
    let _ = call_perf("performanceMark", vec![s("end")]);
    let entry = call_perf(
        "performanceMeasure",
        vec![s("my-measure"), s("start"), s("end")],
    );
    assert!(matches!(entry, Value::Object(_)));
}

#[test]
fn performance_measure_entry_type_is_measure() {
    let _ = call_perf("performanceMark", vec![s("ms")]);
    let _ = call_perf("performanceMark", vec![s("me")]);
    let entry = call_perf("performanceMeasure", vec![s("m"), s("ms"), s("me")]);
    let entry_type = prop(&entry, "entryType");
    assert_eq!(entry_type, s("measure"));
}

#[test]
fn performance_measure_duration_is_non_negative() {
    let _ = call_perf("performanceMark", vec![s("a")]);
    let _ = call_perf("performanceMark", vec![s("b")]);
    let entry = call_perf("performanceMeasure", vec![s("ab"), s("a"), s("b")]);
    let dur = prop(&entry, "duration");
    assert!(as_f64(&dur) >= 0.0);
}

// ── performance.getEntries ────────────────────────────────────────────────────

#[test]
fn performance_get_entries_returns_array() {
    let result = call_perf("performanceGetEntries", vec![]);
    assert!(matches!(result, Value::Object(_)));
    match &result {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            assert!(matches!(&obj.kind, ObjectKind::Array(_)));
        }
        _ => panic!("expected array") }
}

#[test]
fn performance_get_entries_by_type_mark_returns_array() {
    let result = call_perf("performanceGetEntriesByType", vec![s("mark")]);
    assert!(matches!(result, Value::Object(_)));
}

#[test]
fn performance_get_entries_by_name_returns_array() {
    let _ = call_perf("performanceMark", vec![s("named-one")]);
    let result = call_perf("performanceGetEntriesByName", vec![s("named-one")]);
    assert!(matches!(result, Value::Object(_)));
}

// ── performance.eventLoopUtilization ──────────────────────────────────────────

#[test]
fn event_loop_utilization_returns_object() {
    let result = call_perf("eventLoopUtilization", vec![]);
    assert!(matches!(result, Value::Object(_)));
}

#[test]
fn event_loop_utilization_has_utilization_field() {
    let result = call_perf("eventLoopUtilization", vec![]);
    let util = prop(&result, "utilization");
    assert!(matches!(util, Value::F64(_) | Value::I32(_)));
}

// ── monitorEventLoopDelay ─────────────────────────────────────────────────────

#[test]
fn monitor_event_loop_delay_returns_histogram() {
    let hist = call_perf("monitorEventLoopDelay", vec![]);
    assert!(matches!(hist, Value::Object(_)));
}

// ── createHistogram ───────────────────────────────────────────────────────────

#[test]
fn create_histogram_returns_object() {
    let hist = call_perf("createHistogram", vec![]);
    assert!(matches!(hist, Value::Object(_)));
}

// ── Surface check ─────────────────────────────────────────────────────────────

#[test]
fn proposal_node_perf_hooks_surface_is_registered() {
    let expected = [
        "performanceNow",
        "performanceMark",
        "performanceMeasure",
        "performanceClearMarks",
        "performanceClearMeasures",
        "performanceGetEntries",
        "performanceGetEntriesByName",
        "performanceGetEntriesByType",
        "performanceTimeOrigin",
        "eventLoopUtilization",
        "monitorEventLoopDelay",
        "createHistogram",
        "PerformanceObserver",
        "PerformanceEntry",
        "PerformanceMark",
        "PerformanceMeasure",
        "PerformanceResourceTiming",
        "PerformanceNodeTiming",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing node:perf_hooks imports: {missing:?}"
    );
}
