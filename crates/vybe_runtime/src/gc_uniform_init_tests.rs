use super::{allocate_array_values, reserve_array_values};

#[test]
fn uniform_boolean_and_undefined_fills_preserve_exact_variants() {
    for value in [Value::Undefined, Value::Bool(false), Value::Bool(true)] {
        for len in [0usize, 1, 2, 7, 1024, 64000] {
            let values = allocate_array_values(len, value.clone()).unwrap();
            assert_eq!(values.len(), len);
            for actual in &values {
                match (actual, &value) {
                    (Value::Bool(actual), Value::Bool(expected)) => assert_eq!(actual, expected),
                    (Value::Undefined, Value::Undefined) => {},
                    _ => panic!("uniform fill changed the inline value variant"),
                }
            }
        }
        assert!(allocate_array_values(usize::MAX, value).is_err());
    }
}

#[test]
#[ignore = "paired inline boolean/undefined allocation timing"]
fn uniform_boolean_and_undefined_initialization_benchmark() {
    for (label, value) in [
        ("false", Value::Bool(false)), ("true", Value::Bool(true)),
        ("undefined", Value::Undefined),
    ] {
        for len in [0usize, 1, 16, 1024, 64000] {
            let rounds = (65536 / len.max(1)).clamp(4, 256);
            let mut old = Vec::new();
            let mut new = Vec::new();
            for trial in 0..15 {
                if trial % 2 == 0 {
                    old.push(measure(previous, len, &value, rounds));
                    new.push(measure(allocate_array_values, len, &value, rounds));
                } else {
                    new.push(measure(allocate_array_values, len, &value, rounds));
                    old.push(measure(previous, len, &value, rounds));
                }
            }
            old.sort();
            new.sort();
            println!("inline-fill value={label} len={len} previous_us={:.3} current_us={:.3} ratio={:.3}",
                old[7].as_secs_f64() * 1e6, new[7].as_secs_f64() * 1e6,
                old[7].as_secs_f64() / new[7].as_secs_f64());
        }
    }
}

#[test]
fn uniform_typed_null_fills_preserve_native_registry_ids() {
    for id in [0usize, 1, 127, 128, 65535, usize::MAX] {
        for len in [0, 1, 2, 7, 1024, 64000] {
            let values = allocate_array_values(len, Value::TypedNull(id)).unwrap();
            assert_eq!(values.len(), len);
            for value in &values {
                assert!(matches!(value, Value::TypedNull(actual) if *actual == id));
            }
        }
    }
    assert!(allocate_array_values(usize::MAX, Value::TypedNull(1)).is_err());
}

#[test]
#[ignore = "paired allocation timing benchmark"]
fn uniform_typed_null_initialization_benchmark() {
    for id in [0usize, 1, 65535] {
        let value = Value::TypedNull(id);
        for len in [0usize, 1, 16, 1024, 64000] {
            let rounds = (65536 / len.max(1)).clamp(4, 256);
            let mut old = Vec::new();
            let mut new = Vec::new();
            for trial in 0..15 {
                if trial % 2 == 0 {
                    old.push(measure(previous, len, &value, rounds));
                    new.push(measure(allocate_array_values, len, &value, rounds));
                } else {
                    new.push(measure(allocate_array_values, len, &value, rounds));
                    old.push(measure(previous, len, &value, rounds));
                }
            }
            old.sort();
            new.sort();
            println!("typed-null id={id} len={len} previous_us={:.3} current_us={:.3} ratio={:.3}",
                old[7].as_secs_f64() * 1e6, new[7].as_secs_f64() * 1e6,
                old[7].as_secs_f64() / new[7].as_secs_f64());
        }
    }
}
use crate::value::Object;
use crate::{Value, VMError};
use std::hint::black_box;
use std::sync::Arc;
use std::time::{Duration, Instant};

fn assert_payload(actual: &Value, expected: &Value) {
    match (actual, expected) {
        (Value::I32(a), Value::I32(b)) => assert_eq!(a, b),
        (Value::I64(a), Value::I64(b)) => assert_eq!(a, b),
        (Value::F32(a), Value::F32(b)) => assert_eq!(a.to_bits(), b.to_bits()),
        (Value::F64(a), Value::F64(b)) => assert_eq!(a.to_bits(), b.to_bits()),
        (Value::V128(a), Value::V128(b)) => assert_eq!(a, b),
        (Value::Null, Value::Null) => (),
        _ => panic!("uniform initialization changed the element variant"),
    }
}

#[test]
fn uniform_numeric_fills_preserve_variants_and_every_payload_bit() {
    for value in [
        Value::I32(-1), Value::I64(i64::MIN),
        Value::F32(-0.0), Value::F32(f32::from_bits(0x7f800001)),
        Value::F64(-0.0), Value::F64(f64::from_bits(0x7ff0000000000001)),
        Value::V128([0x91; 16]), Value::Null,
    ] {
        for len in [0, 1, 2, 7, 1024, 64000] {
            let values = allocate_array_values(len, value.clone()).unwrap();
            assert_eq!(values.len(), len);
            for actual in &values { assert_payload(actual, &value); }
        }
    }
}

#[test]
fn uniform_reference_fills_preserve_identity_and_drop_counts() {
    let object = crate::heap::alloc(Object::new());
    for len in [0, 1, 7, 1024] {
        let values = allocate_array_values(len, Value::Object(object.clone())).unwrap();
        assert_eq!(values.len(), len);
        assert_eq!(Arc::strong_count(&object), len + 1);
        for value in &values {
            let Value::Object(value) = value else { panic!("reference changed variant") };
            assert!(Arc::ptr_eq(value, &object));
        }
        drop(values);
        assert_eq!(Arc::strong_count(&object), 1);
    }
    assert!(allocate_array_values(usize::MAX, Value::Object(object.clone())).is_err());
    assert_eq!(Arc::strong_count(&object), 1);
}

fn previous(len: usize, value: Value) -> Result<Vec<Value>, VMError> {
    let mut values = reserve_array_values(len)?;
    values.resize(len, value);
    Ok(values)
}

fn measure(
    kernel: fn(usize, Value) -> Result<Vec<Value>, VMError>,
    len: usize, source: &Value, rounds: usize,
) -> Duration {
    let mut elapsed = Duration::ZERO;
    for _ in 0..rounds {
        let value = source.clone();
        let start = Instant::now();
        let values = kernel(black_box(len), black_box(value)).unwrap();
        elapsed += start.elapsed();
        black_box(&values);
        drop(values); // Do not include destruction in the allocation/init timer.
    }
    elapsed / rounds as u32
}

#[test]
#[ignore = "paired allocation/init timing experiment"]
fn uniform_array_initialization_benchmark() {
    for (name, value) in [
        ("i32", Value::I32(0)), ("i64", Value::I64(0)),
        ("f32", Value::F32(0.0)), ("f64", Value::F64(0.0)),
        ("v128", Value::V128([0; 16])), ("null", Value::Null),
        ("ref", Value::Object(crate::heap::alloc(Object::new()))),
    ] {
        for len in [0usize, 1, 16, 1024, 64000] {
            let rounds = (65536 / len.max(1)).clamp(4, 256);
            let mut old = Vec::new();
            let mut new = Vec::new();
            for trial in 0..15 {
                if trial % 2 == 0 {
                    old.push(measure(previous, len, &value, rounds));
                    new.push(measure(allocate_array_values, len, &value, rounds));
                } else {
                    new.push(measure(allocate_array_values, len, &value, rounds));
                    old.push(measure(previous, len, &value, rounds));
                }
            }
            old.sort_unstable(); new.sort_unstable();
            let before = old[7].as_secs_f64() * 1e6;
            let after = new[7].as_secs_f64() * 1e6;
            println!("{name:4} len={len:5} previous_us={before:.3} current_us={after:.3} ratio={:.3}", before / after);
        }
    }
}
