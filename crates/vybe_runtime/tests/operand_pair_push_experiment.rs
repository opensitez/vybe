//! Safe pair-push experiment; no production stack policy change.

use std::sync::Arc;
use vybe_runtime::Value;

#[inline(never)]
fn previous(stack: &mut Vec<Value>, first: Value, second: Value) {
    stack.push(first);
    stack.push(second);
}

#[inline(never)]
fn candidate(stack: &mut Vec<Value>, first: Value, second: Value) {
    stack.extend([first, second]);
}

fn assert_same(actual: &Value, expected: &Value) {
    match (actual, expected) {
        (Value::Null, Value::Null) => {}
        (Value::I32(a), Value::I32(b)) => assert_eq!(a, b),
        (Value::I64(a), Value::I64(b)) => assert_eq!(a, b),
        (Value::F32(a), Value::F32(b)) => assert_eq!(a.to_bits(), b.to_bits()),
        (Value::F64(a), Value::F64(b)) => assert_eq!(a.to_bits(), b.to_bits()),
        (Value::V128(a), Value::V128(b)) => assert_eq!(a, b),
        (Value::String(a), Value::String(b)) => assert!(Arc::ptr_eq(a, b)),
        _ => panic!("pair push changed a variant"),
    }
}

#[test]
fn pair_push_preserves_order_bits_and_references_through_growth() {
    let values = [
        Value::Null,
        Value::I32(i32::MIN),
        Value::I64(i64::MAX),
        Value::F32(f32::from_bits(0x7f80_0001)),
        Value::F64(f64::from_bits(0x8000_0000_0000_0000)),
        Value::V128([0xa5; 16]),
        Value::String(String::from("pair identity").into()),
    ];
    for capacity in [0, 1, 2, 3, 16] {
        for first in &values {
            for second in &values {
                let mut old = Vec::with_capacity(capacity);
                let mut new = Vec::with_capacity(capacity);
                old.push(Value::I32(19));
                new.push(Value::I32(19));
                previous(&mut old, first.clone(), second.clone());
                candidate(&mut new, first.clone(), second.clone());
                assert_eq!(new.len(), old.len());
                for (actual, expected) in new.iter().zip(&old) {
                    assert_same(actual, expected);
                }
            }
        }
    }
}

#[test]
fn pair_push_does_not_retain_references_after_stack_clear() {
    for push in [previous, candidate] {
        let value = Value::String(String::from("release").into());
        let weak = match &value {
            Value::String(text) => Arc::downgrade(text),
            _ => unreachable!(),
        };
        let mut stack = Vec::new();
        push(&mut stack, value.clone(), value);
        assert!(weak.upgrade().is_some());
        stack.clear();
        assert!(weak.upgrade().is_none());
    }
}

#[test]
#[ignore = "paired safe operand pair-push experiment; not VM execution timing"]
fn paired_pair_push_benchmark() {
    use std::hint::black_box;
    use std::time::Instant;
    type Pusher = fn(&mut Vec<Value>, Value, Value);
    fn sample(push: Pusher, stack: &mut Vec<Value>, pairs: usize) -> f64 {
        stack.clear();
        let start = Instant::now();
        for index in 0..pairs {
            black_box(push)(
                black_box(stack),
                Value::I32(index as i32),
                Value::I32(!(index as i32)),
            );
        }
        let elapsed = start.elapsed().as_secs_f64();
        assert_eq!(stack.len(), pairs * 2);
        for (index, pair) in stack.chunks_exact(2).enumerate() {
            assert!(matches!(pair[0], Value::I32(value) if value == index as i32));
            assert!(matches!(pair[1], Value::I32(value) if value == !(index as i32)));
        }
        elapsed
    }
    for pairs in [4096usize, 65_536, 262_144] {
        let mut stack = Vec::with_capacity(pairs * 2);
        sample(previous, &mut stack, pairs);
        sample(candidate, &mut stack, pairs);
        let mut ratios = Vec::with_capacity(31);
        for pair in 0..31 {
            let mut times = [0.0; 2];
            for is_new in if pair % 2 == 0 { [false, true] } else { [true, false] } {
                times[usize::from(is_new)] = sample(if is_new { candidate } else { previous }, &mut stack, pairs);
            }
            assert!(times.iter().all(|time| time.is_finite() && *time > 0.0));
            ratios.push(times[1] / times[0]);
        }
        ratios.sort_by(f64::total_cmp);
        println!("safe pair push pairs={pairs} extend/push={:.6}; preallocated helper only, validation and clear excluded", ratios[15]);
    }
}
