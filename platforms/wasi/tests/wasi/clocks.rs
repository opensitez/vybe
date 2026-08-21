//! Behaviour tests for WASI clocks proposal interfaces.
//! Only real WASI interfaces: wasi:clocks/wall-clock, wasi:clocks/monotonic-clock.

use std::time::{Duration, SystemTime, UNIX_EPOCH};

use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::{Object, Value};
use vybe_runtime::{Chunk, Op, VM};

static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn invoke(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-clocks-test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    for value in args {
        match value {
            Value::I32(n) => chunk.emit_i32_const(n, 0),
            Value::I64(n) => chunk.emit_i64_const(n, 0),
            Value::F32(f) => chunk.emit_f32_const(f, 0),
            Value::F64(f) => chunk.emit_f64_const(f, 0),
            Value::Bool(b) => chunk.emit_bool_const(b, 0),
            Value::String(text) => chunk.emit_string_const(&text, 0),
            Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0),
            other => {
                let global_name = format!(
                    "__test_arg_{}",
                    TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
                );
                vm.set_global_owned(global_name.clone(), other);
                let ci = chunk.intern_string_constant(&global_name);
                chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
            }
        }
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    vm.run(vec![chunk]).expect("VM run failed")
}

fn invoke_result(module: &str, name: &str, args: Vec<Value>) -> Result<Value, String> {
    let mut chunk = Chunk::new("<wasi-clocks-test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    for value in args {
        match value {
            Value::I32(n) => chunk.emit_i32_const(n, 0),
            Value::I64(n) => chunk.emit_i64_const(n, 0),
            Value::F32(f) => chunk.emit_f32_const(f, 0),
            Value::F64(f) => chunk.emit_f64_const(f, 0),
            Value::Bool(b) => chunk.emit_bool_const(b, 0),
            Value::String(text) => chunk.emit_string_const(&text, 0),
            Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0),
            other => panic!("no spec const emitter for test argument {other:?}"),
        }
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).map_err(|error| error.message)
}

fn has_import(module: &str, name: &str) -> bool {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
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

// ── wasi:clocks/wall-clock ──────────────────────────────────────────

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

// ── wasi:clocks/monotonic-clock ─────────────────────────────────────

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
fn monotonic_clock_resolution_is_exactly_one_nanosecond() {
    let resolution = invoke("wasi:clocks/monotonic-clock", "resolution", vec![]);
    assert_eq!(resolution, Value::F64(1.0));
}

// ── wasi:clocks/monotonic-clock proposal surface ────────────────────

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

// ── wasi:clocks/timezone ────────────────────────────────────────────────────

/// New York is UTC−5 in January and UTC−4 in July.
///
/// One assertion covers the whole TZif reader at once: selecting the 64-bit
/// data block over the v1 one, searching the transition list, and reporting the
/// daylight-saving flag. The reader is what `utc-offset` and `iana-id` answer
/// from, so a parse regression shows up here rather than as a silently wrong
/// timestamp somewhere downstream.
#[test]
fn tzif_reader_finds_both_sides_of_a_dst_transition() {
    let Ok(bytes) = std::fs::read("/usr/share/zoneinfo/America/New_York") else {
        eprintln!("no system tzdb on this host; skipping");
        return;
    };

    // 2021-01-01T00:00:00Z and 2021-07-01T00:00:00Z.
    let january = vybe_platform_wasi::clock::tzif_type_at_bytes(&bytes, 1_609_459_200);
    let july = vybe_platform_wasi::clock::tzif_type_at_bytes(&bytes, 1_625_097_600);

    assert_eq!(january, Some((-5 * 3600, false)), "EST is UTC-5, not DST");
    assert_eq!(july, Some((-4 * 3600, true)), "EDT is UTC-4, in DST");
}

/// `utc-offset` answers `option<s64>` NANOSECONDS.
///
/// The older interface returned `s32` SECONDS and always produced a value; a
/// host that cannot determine its zone now answers nothing, because `0` would
/// claim UTC for every machine that simply has no tzdb.
#[test]
fn utc_offset_is_nanoseconds_or_absent() {
    let mut instant = Object::new();
    instant
        .properties
        .insert("seconds".into(), Value::F64(1_609_459_200.0));
    instant
        .properties
        .insert("nanoseconds".into(), Value::F64(0.0));
    let when = Value::Object(vybe_runtime::heap::alloc(instant));

    match invoke("wasi:clocks/timezone", "utc-offset", vec![when]) {
        Value::Null => {} // no zone configured on this host — permitted
        other => {
            let nanos = other.as_f64();
            assert_eq!(
                nanos % 1_000_000_000.0,
                0.0,
                "a whole-second offset must be a whole number of nanoseconds, got {nanos}"
            );
            assert!(
                nanos.abs() < 86_400_000_000_000.0,
                "the interface bounds the magnitude below one day, got {nanos}"
            );
        }
    }
}

/// `iana-id` is `option<string>` — a zone name or nothing, never `""`.
#[test]
fn iana_id_is_a_zone_name_or_nothing() {
    match invoke("wasi:clocks/timezone", "iana-id", vec![]) {
        Value::Null => {}
        Value::String(id) => assert!(!id.is_empty(), "an empty id must be reported as none"),
        other => panic!("iana-id must be option<string>, got {other:?}"),
    }
}

/// `system-clock` is the current name for what used to be `wall-clock`; both
/// are bound and must agree, since they are one clock.
#[test]
fn system_clock_and_wall_clock_are_the_same_clock() {
    let system = invoke("wasi:clocks/system-clock", "now", vec![]);
    let wall = invoke("wasi:clocks/wall-clock", "now", vec![]);
    let seconds_of = |value: &Value| prop(value, "seconds").as_f64();
    assert!(
        (seconds_of(&system) - seconds_of(&wall)).abs() <= 1.0,
        "system-clock and wall-clock must report the same moment"
    );
    assert_eq!(
        invoke("wasi:clocks/system-clock", "get-resolution", vec![]),
        Value::F64(1.0),
        "get-resolution answers a duration in nanoseconds, not a record"
    );
}
