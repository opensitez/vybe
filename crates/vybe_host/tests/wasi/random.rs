use vybe_bytecode::value::{ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn invoke(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-random-test>");
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
    let mut chunk = Chunk::new("<wasi-random-test>");
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

fn array_len(value: &Value) -> usize {
    let Value::Object(object) = value else { return 0 };
    let object = object.lock().unwrap();
    let ObjectKind::Array(values) = &object.kind else { return 0 };
    values.len()
}

fn byte_values(value: &Value) -> Vec<i32> {
    let Value::Object(object) = value else { return Vec::new() };
    let object = object.lock().unwrap();
    let ObjectKind::Array(values) = &object.kind else { return Vec::new() };
    values.iter().map(|value| value.as_i32()).collect()
}

#[test]
fn get_random_bytes_returns_requested_length() {
    let result = invoke("wasi:random/random", "get-random-bytes", vec![Value::F64(16.0)]);
    assert_eq!(array_len(&result), 16);
    assert!(byte_values(&result).iter().all(|value| (0..=255).contains(value)));
}

#[test]
fn get_random_bytes_zero_length_returns_empty_array() {
    let result = invoke("wasi:random/random", "get-random-bytes", vec![Value::F64(0.0)]);
    assert_eq!(array_len(&result), 0);
}

#[test]
fn get_random_bytes_defaults_to_empty_when_length_missing() {
    let result = invoke("wasi:random/random", "get-random-bytes", vec![]);
    assert_eq!(array_len(&result), 0);
}

#[test]
fn get_random_bytes_negative_length_clamps_to_empty_array() {
    let result = invoke("wasi:random/random", "get-random-bytes", vec![Value::F64(-1.0)]);
    assert_eq!(array_len(&result), 0);
}

#[test]
fn get_random_u64_returns_numeric_value() {
    let result = invoke("wasi:random/random", "get-random-u64", vec![]);
    assert!(matches!(result, Value::F64(_)));
}

#[test]
fn get_insecure_random_bytes_returns_requested_length() {
    let result = invoke(
        "wasi:random/insecure",
        "get-insecure-random-bytes",
        vec![Value::F64(8.0)],
    );
    assert_eq!(array_len(&result), 8);
}

#[test]
fn get_insecure_random_bytes_zero_length_returns_empty_array() {
    let result = invoke(
        "wasi:random/insecure",
        "get-insecure-random-bytes",
        vec![Value::F64(0.0)],
    );
    assert_eq!(array_len(&result), 0);
}

#[test]
fn get_insecure_random_bytes_defaults_to_empty_when_length_missing() {
    let result = invoke("wasi:random/insecure", "get-insecure-random-bytes", vec![]);
    assert_eq!(array_len(&result), 0);
}

#[test]
fn get_insecure_random_bytes_negative_length_clamps_to_empty_array() {
    let result = invoke(
        "wasi:random/insecure",
        "get-insecure-random-bytes",
        vec![Value::F64(-1.0)],
    );
    assert_eq!(array_len(&result), 0);
}

#[test]
fn get_insecure_random_bytes_large_length_returns_requested_length() {
    let result = invoke(
        "wasi:random/insecure",
        "get-insecure-random-bytes",
        vec![Value::F64(256.0)],
    );
    assert_eq!(array_len(&result), 256);
    assert!(byte_values(&result).iter().all(|value| (0..=255).contains(value)));
}

#[test]
fn get_insecure_random_bytes_values_stay_in_byte_range() {
    let result = invoke(
        "wasi:random/insecure",
        "get-insecure-random-bytes",
        vec![Value::F64(32.0)],
    );
    assert!(byte_values(&result).iter().all(|value| (0..=255).contains(value)));
}

#[test]
fn get_insecure_random_u64_returns_numeric_value() {
    let result = invoke("wasi:random/insecure", "get-insecure-random-u64", vec![]);
    assert!(matches!(result, Value::F64(_)));
}

#[test]
fn insecure_seed_returns_two_number_tuple() {
    let result = invoke("wasi:random/insecure-seed", "insecure-seed", vec![]);
    assert_eq!(array_len(&result), 2);
    if let Value::Object(object) = &result {
        let object = object.lock().unwrap();
        if let ObjectKind::Array(values) = &object.kind {
            assert!(matches!(values.first(), Some(Value::F64(_))));
            assert!(matches!(values.get(1), Some(Value::F64(_))));
            return;
        }
    }
    panic!("insecure-seed should return [u64, u64], got {:?}", result);
}

#[test]
fn insecure_seed_values_are_non_negative_numbers() {
    let result = invoke("wasi:random/insecure-seed", "insecure-seed", vec![]);
    let Value::Object(object) = result else {
        panic!("insecure-seed should return array");
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(values) = &object.kind else {
        panic!("insecure-seed should return array");
    };
    assert_eq!(values.len(), 2);
    for value in values {
        let Value::F64(number) = value else {
            panic!("seed element should be numeric");
        };
        assert!(*number >= 0.0);
    }
}

#[test]
fn get_random_bytes_large_length_returns_requested_length() {
    let result = invoke("wasi:random/random", "get-random-bytes", vec![Value::F64(256.0)]);
    assert_eq!(array_len(&result), 256);
    assert!(byte_values(&result).iter().all(|value| (0..=255).contains(value)));
}

#[test]
fn random_helper_returns_unit_interval_value() {
    let result = invoke("wasi:random/random", "random", vec![]);
    let Value::F64(number) = result else {
        panic!("random should return f64");
    };
    assert!((0.0..1.0).contains(&number));
}

#[test]
fn random_helper_stays_in_unit_interval_across_multiple_samples() {
    for _ in 0..64 {
        let result = invoke("wasi:random/random", "random", vec![]);
        let Value::F64(number) = result else {
            panic!("random should return f64");
        };
        assert!((0.0..1.0).contains(&number));
    }
}

#[test]
fn random_int_returns_bound_when_range_collapses() {
    let result = invoke(
        "wasi:random/random",
        "randomInt",
        vec![Value::F64(5.0), Value::F64(5.0)],
    );
    assert_eq!(result, Value::F64(5.0));
}

#[test]
fn random_int_returns_min_when_max_is_less_than_min() {
    let result = invoke(
        "wasi:random/random",
        "randomInt",
        vec![Value::F64(8.0), Value::F64(3.0)],
    );
    assert_eq!(result, Value::F64(8.0));
}

#[test]
fn random_int_defaults_to_zero_through_hundred_when_arguments_missing() {
    for _ in 0..32 {
        let result = invoke("wasi:random/random", "randomInt", vec![]);
        let Value::F64(number) = result else {
            panic!("randomInt should return f64");
        };
        assert!((0.0..=100.0).contains(&number));
    }
}

#[test]
fn random_int_supports_negative_ranges() {
    for _ in 0..32 {
        let result = invoke(
            "wasi:random/random",
            "randomInt",
            vec![Value::F64(-3.0), Value::F64(2.0)],
        );
        let Value::F64(number) = result else {
            panic!("randomInt should return f64");
        };
        assert!((-3.0..=2.0).contains(&number));
    }
}

#[test]
fn random_int_stays_within_inclusive_range() {
    for _ in 0..32 {
        let result = invoke(
            "wasi:random/random",
            "randomInt",
            vec![Value::F64(3.0), Value::F64(7.0)],
        );
        let Value::F64(number) = result else {
            panic!("randomInt should return f64");
        };
        assert!((3.0..=7.0).contains(&number));
    }
}

#[test]
fn uuid_returns_v4_shaped_identifier() {
    let result = invoke("wasi:random/random", "uuid", vec![]);
    let Value::String(text) = result else {
        panic!("uuid should return string");
    };
    let text = text.as_ref();
    assert_eq!(text.len(), 36);
    assert_eq!(&text[8..9], "-");
    assert_eq!(&text[13..14], "-");
    assert_eq!(&text[18..19], "-");
    assert_eq!(&text[23..24], "-");
    assert_eq!(&text[14..15], "4");
    assert!(matches!(&text[19..20], "8" | "9" | "a" | "b"));
}

#[test]
fn uuid_calls_return_distinct_values() {
    let first = invoke("wasi:random/random", "uuid", vec![]);
    let second = invoke("wasi:random/random", "uuid", vec![]);
    assert_ne!(first, second);
}

#[test]
fn proposal_insecure_seed_get_insecure_seed_import_resolves() {
    assert!(
        invoke_result("wasi:random/insecure-seed", "get-insecure-seed", vec![]).is_ok(),
        "wasi:random/insecure-seed.get-insecure-seed should be covered by the random category"
    );
}