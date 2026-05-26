use std::sync::{Arc, Mutex};
use vybe_bytecode::value::Object;
use vybe_bytecode::value::Value;
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<ecma-math-test>");
    let import_idx = chunk.add_import("ecma:math", name);
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

fn array(values: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(values))))
}

#[test]
fn floor_and_ceil_round_toward_expected_bounds() {
    assert_eq!(invoke("floor", vec![Value::F64(3.9)]), Value::F64(3.0));
    assert_eq!(invoke("ceil", vec![Value::F64(3.1)]), Value::F64(4.0));
}

#[test]
fn sign_reports_negative_zero_and_positive_values() {
    assert_eq!(invoke("sign", vec![Value::F64(-4.0)]), Value::F64(-1.0));
    assert_eq!(invoke("sign", vec![Value::F64(0.0)]), Value::F64(0.0));
    assert_eq!(invoke("sign", vec![Value::F64(8.0)]), Value::F64(1.0));
}

#[test]
fn log_with_explicit_base_uses_change_of_base() {
    assert_eq!(invoke("log", vec![Value::F64(100.0), Value::F64(10.0)]), Value::F64(2.0));
}

#[test]
fn clamp_limits_to_closed_interval() {
    assert_eq!(
        invoke("clamp", vec![Value::F64(10.0), Value::F64(0.0), Value::F64(5.0)]),
        Value::F64(5.0)
    );
    assert_eq!(
        invoke("clamp", vec![Value::F64(-2.0), Value::F64(0.0), Value::F64(5.0)]),
        Value::F64(0.0)
    );
}

#[test]
fn imul_multiplies_using_32_bit_wrapping_rules() {
    assert_eq!(invoke("imul", vec![Value::F64(3.0), Value::F64(4.0)]), Value::F64(12.0));
}

#[test]
fn random_and_rnd_return_unit_interval_values() {
    for name in ["random", "rnd"] {
        let Value::F64(number) = invoke(name, vec![]) else {
            panic!("{} should return f64", name);
        };
        assert!((0.0..1.0).contains(&number), "{} out of range: {}", name, number);
    }
}

#[test]
fn randomize_returns_null() {
    assert!(matches!(invoke("randomize", vec![]), Value::Null));
}

#[test]
fn min_of_and_max_of_materialize_iterables() {
    let values = array(vec![Value::F64(3.0), Value::F64(1.0), Value::F64(4.0)]);
    assert_eq!(invoke("minOf", vec![values.clone()]), Value::F64(1.0));
    assert_eq!(invoke("maxOf", vec![values]), Value::F64(4.0));
}

#[test]
fn sum_precise_handles_catastrophic_cancellation_case() {
    let values = array(vec![Value::F64(1.0e16), Value::F64(1.0), Value::F64(-1.0e16)]);
    assert_eq!(invoke("sumPrecise", vec![values]), Value::F64(1.0));
}

#[test]
fn cbrt_and_hypot_follow_standard_library_results() {
    assert_eq!(invoke("cbrt", vec![Value::F64(27.0)]), Value::F64(3.0));
    assert_eq!(invoke("hypot", vec![Value::F64(3.0), Value::F64(4.0)]), Value::F64(5.0));
}