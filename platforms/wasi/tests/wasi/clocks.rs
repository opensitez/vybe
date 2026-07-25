//! Behaviour tests for WASI clocks proposal interfaces.
//! Only real WASI interfaces: wasi:clocks/wall-clock, wasi:clocks/monotonic-clock.

use std::time::{Duration, SystemTime, UNIX_EPOCH};

use vybe_bytecode::value::Value;
use vybe_bytecode::{Chunk, Op, VM};
use vybe_bytecode::capabilities::Capabilities;
use vybe_emitter::platforms::register_platforms;

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
    register_platforms(&mut vm, &Capabilities::all());
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
