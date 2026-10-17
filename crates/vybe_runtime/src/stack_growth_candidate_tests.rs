use crate::Value;
use std::{hint::black_box, sync::Arc, time::Instant};

fn reserved_push(values: &mut Vec<Value>, len: usize) {
    if len <= values.len() {
        values.truncate(len);
        return;
    }
    values.reserve(len - values.len());
    while values.len() < len {
        values.push(Value::Null);
    }
}

fn resize_with(values: &mut Vec<Value>, len: usize) {
    values.resize_with(len, || Value::Null);
}

#[test]
fn growth_candidates_preserve_prefix_and_reference_lifetimes() {
    let text: Arc<str> = Arc::from("live prefix");
    for grow in [reserved_push as fn(&mut Vec<Value>, usize), resize_with] {
        let mut values = vec![Value::String(text.clone()), Value::Undefined];
        grow(&mut values, 199);
        assert_eq!(Arc::strong_count(&text), 2);
        match &values[0] {
            Value::String(actual) => assert!(Arc::ptr_eq(actual, &text)),
            _ => panic!("prefix changed"),
        }
        assert!(matches!(values[1], Value::Undefined));
        assert!(values[2..].iter().all(|value| matches!(value, Value::Null)));
        grow(&mut values, 0);
        assert!(values.is_empty());
        assert_eq!(Arc::strong_count(&text), 1);
    }
}

#[test]
#[ignore = "paired safe stack growth candidates; run explicitly"]
fn stack_growth_candidates_benchmark() {
    let methods: [fn(&mut Vec<Value>, usize); 3] = [
        super::resize_null_values, reserved_push, resize_with,
    ];
    for len in [16usize, 199, 1024, 64000] {
        let mut buffers: [Vec<Value>; 3] = std::array::from_fn(|_| Vec::with_capacity(len));
        let mut samples: [Vec<f64>; 3] = std::array::from_fn(|_| Vec::new());
        let iterations = (65536 / len).clamp(1, 256);
        for trial in 0..21 {
            let mut nanos = [0u128; 3];
            for _ in 0..iterations {
                for order in 0..3 {
                    let which = (trial + order) % 3;
                    buffers[which].clear();
                    let start = Instant::now();
                    methods[which](black_box(&mut buffers[which]), black_box(len));
                    nanos[which] += start.elapsed().as_nanos();
                    assert_eq!(buffers[which].len(), len);
                }
            }
            for which in 0..3 { samples[which].push(nanos[which] as f64 / iterations as f64); }
        }
        for values in &mut samples { values.sort_by(f64::total_cmp); }
        println!("safe-growth len={len} current_ns={:.2} reserved_push_ns={:.2} resize_with_ns={:.2}",
            samples[0][10], samples[1][10], samples[2][10]);
    }
}
