use super::{resize_undefined_values, resize_null_values};
use crate::{Value, value::Object};
use std::{hint::black_box, sync::Arc, time::Instant};

#[test]
fn stack_growth_and_shrink_preserve_prefix_identity_and_drop_counts() {
    for start in [0usize, 1, 7, 1024] {
        for len in [0usize, 1, 7, 16, 1024, 64000] {
            let object = crate::heap::alloc(Object::new());
            let mut values: Vec<Value> = (0..start).map(|_| Value::Object(object.clone())).collect();
            resize_undefined_values(&mut values, len);
            assert_eq!(values.len(), len);
            for value in &values[..start.min(len)] {
                assert!(matches!(value, Value::Object(actual) if Arc::ptr_eq(actual, &object)));
            }
            assert!(values[start.min(len)..].iter().all(|v| matches!(v, Value::Undefined)));
            assert_eq!(Arc::strong_count(&object), start.min(len) + 1);
            drop(values);
            assert_eq!(Arc::strong_count(&object), 1);
        }
    }
}

#[test]
fn stack_growth_keeps_numeric_bits_native_type_ids_and_vector_bytes() {
    let mut values = vec![Value::F32(f32::from_bits(0x7f800001)),
        Value::F64(f64::from_bits(0x7ff0000000000001)), Value::TypedNull(usize::MAX),
        Value::V128([0x91; 16]), Value::Bool(false), Value::Null];
    resize_undefined_values(&mut values, 64);
    assert!(matches!(values[0], Value::F32(v) if v.to_bits() == 0x7f800001));
    assert!(matches!(values[1], Value::F64(v) if v.to_bits() == 0x7ff0000000000001));
    assert!(matches!(values[2], Value::TypedNull(usize::MAX)));
    assert!(matches!(values[3], Value::V128(v) if v == [0x91; 16]));
    assert!(matches!(values[4], Value::Bool(false)));
    assert!(matches!(values[5], Value::Null));
    assert!(values[6..].iter().all(|v| matches!(v, Value::Undefined)));
    resize_undefined_values(&mut values, 6);
    assert_eq!(values.len(), 6);
}

fn previous(values: &mut Vec<Value>, len: usize) { values.resize(len, Value::Undefined); }

fn previous_null(values: &mut Vec<Value>, len: usize) { values.resize(len, Value::Null); }

#[test]
fn null_stack_growth_preserves_variant_prefix_and_reference_destruction() {
    for start in [0usize, 1, 7, 1024] {
        for len in [0usize, 1, 7, 16, 1024, 64000] {
            let object = crate::heap::alloc(Object::new());
            let mut values: Vec<Value> = (0..start).map(|_| Value::Object(object.clone())).collect();
            resize_null_values(&mut values, len);
            assert_eq!(values.len(), len);
            for value in &values[..start.min(len)] {
                assert!(matches!(value, Value::Object(actual) if Arc::ptr_eq(actual, &object)));
            }
            assert!(values[start.min(len)..].iter().all(|v| matches!(v, Value::Null)));
            assert_eq!(Arc::strong_count(&object), start.min(len) + 1);
            drop(values);
            assert_eq!(Arc::strong_count(&object), 1);
        }
    }
}

#[test]
#[ignore = "paired operand-stack growth benchmark"]
fn undefined_stack_resize_benchmark() {
  for (label, old_kernel, new_kernel) in [
      ("undefined", previous as fn(&mut Vec<Value>, usize), resize_undefined_values as fn(&mut Vec<Value>, usize)),
      ("null", previous_null as fn(&mut Vec<Value>, usize), resize_null_values as fn(&mut Vec<Value>, usize)),
  ] {
    for len in [0usize, 1, 4, 16, 1024, 64000] {
        for reuse in [false, true] {
            let rounds = (65536 / len.max(1)).clamp(4, 1024);
            let measure = |kernel: fn(&mut Vec<Value>, usize)| {
                let mut total = 0u128;
                let mut values = Vec::with_capacity(if reuse { len } else { 0 });
                for _ in 0..rounds {
                    if !reuse { values = Vec::new(); }
                    let start = Instant::now();
                    kernel(black_box(&mut values), black_box(len));
                    total += start.elapsed().as_nanos();
                    black_box(&values);
                    values.clear();
                }
                total as f64 / rounds as f64
            };
            let mut old = Vec::new();
            let mut new = Vec::new();
            for trial in 0..31 {
                if trial % 2 == 0 { old.push(measure(old_kernel)); new.push(measure(new_kernel)); }
                else { new.push(measure(new_kernel)); old.push(measure(old_kernel)); }
            }
            old.sort_by(f64::total_cmp);
            new.sort_by(f64::total_cmp);
            println!("stack fill={label} len={len} reuse={reuse} old_ns={:.2} new_ns={:.2} ratio={:.2}",
                old[15], new[15], old[15] / new[15]);
        }
    }
  }
}
