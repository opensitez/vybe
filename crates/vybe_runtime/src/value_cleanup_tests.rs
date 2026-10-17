//! Safe cleanup experiment: no production stack behavior changes.
use crate::Value;
use std::{hint::black_box, sync::Arc, time::Instant};

fn drain_destructured(values: &mut Vec<Value>) {
    // Forward draining retains Vec's destruction order and panic cleanup.
    for value in values.drain(..) {
        match value {
            Value::Null | Value::Undefined | Value::TypedNull(_)
            | Value::Bool(_) | Value::I32(_) | Value::I64(_)
            | Value::F32(_) | Value::F64(_) | Value::V128(_) => {}
            Value::String(payload) => drop(payload),
            Value::Object(payload) => drop(payload),
            Value::WeakRef(payload) => drop(payload),
            Value::Symbol(payload) => drop(payload),
            Value::BigInt(payload) => drop(payload),
        }
    }
}

#[test]
fn destructured_cleanup_releases_references_and_reuses_capacity() {
    let text: Arc<str> = Arc::from("retained");
    let mut values = Vec::with_capacity(1024);
    values.extend((0..1024).map(|index| match index % 4 {
        0 => Value::String(text.clone()),
        1 => Value::Symbol(text.clone()),
        2 => Value::TypedNull(usize::MAX),
        _ => Value::V128([0x91; 16]),
    }));
    assert_eq!(Arc::strong_count(&text), 513);
    let capacity = values.capacity();
    drain_destructured(&mut values);
    assert!(values.is_empty());
    assert_eq!(values.capacity(), capacity);
    assert_eq!(Arc::strong_count(&text), 1);
    values.push(Value::F64(f64::from_bits(0x7ff0_0000_0000_0001)));
    drain_destructured(&mut values);
    drain_destructured(&mut values);
    assert!(values.is_empty());
}

#[test]
#[ignore = "paired cleanup microbenchmark; run explicitly"]
fn value_cleanup_benchmark() {
    let text: Arc<str> = Arc::from("benchmark reference");
    for len in [16usize, 1024, 64000] {
        for references in [false, true] {
            let source: Vec<Value> = (0..len).map(|index| {
                if references && index % 4 == 0 {
                    Value::String(text.clone())
                } else {
                    Value::Null
                }
            }).collect();
            let mut baseline = Vec::with_capacity(len);
            let mut candidate = Vec::with_capacity(len);
            let mut samples = [Vec::new(), Vec::new()];
            let iterations = (65536 / len).clamp(1, 256);
            for trial in 0..21 {
                let mut nanos = [0u128; 2];
                for _ in 0..iterations {
                    for order in 0..2 {
                        let which = (trial + order) % 2;
                        let values = if which == 0 { &mut baseline } else { &mut candidate };
                        values.extend(source.iter().cloned());
                        let start = Instant::now();
                        if which == 0 { black_box(&mut *values).clear(); }
                        else { drain_destructured(black_box(&mut *values)); }
                        nanos[which] += start.elapsed().as_nanos();
                        assert!(values.is_empty());
                    }
                }
                for which in 0..2 { samples[which].push(nanos[which] as f64 / iterations as f64); }
            }
            for sample in &mut samples { sample.sort_by(f64::total_cmp); }
            let old = samples[0][10];
            let new = samples[1][10];
            println!("cleanup len={len} references={references} clear_ns={old:.2} drain_ns={new:.2} speedup={:.3}", old / new);
            drop(source);
            assert_eq!(Arc::strong_count(&text), 1);
        }
    }
}
