//! Compiler-free paired benchmark for bound-callback decoding.
//! Measures metadata locks, capture snapshot creation, and snapshot destruction,
//! not complete bind/invocation or guest/application execution.

use std::{hint::black_box, time::{Duration, Instant}};
use vybe_runtime::{value::{Object, ObjectKind}, Value};

struct PreviousPrepared {
    _target: Value,
    _receiver: Value,
    _prototype: Value,
    _partials: Vec<Value>,
}

fn previous(callback: &Value) -> Option<PreviousPrepared> {
    let Value::Object(wrapper) = callback else { return None };
    let wrapper = wrapper.lock().unwrap();
    let Some(Value::Object(backing)) = wrapper.properties.get("__bound_args") else {
        return None;
    };
    if !matches!(wrapper.properties.get("name"), Some(Value::String(name)) if name.starts_with("bound ")) {
        return None;
    }
    let backing = backing.lock().unwrap();
    let ObjectKind::Array(values) = &backing.kind else { return None };
    if values.len() < 3 { return None; }
    Some(PreviousPrepared {
        _target: values[0].clone(),
        _receiver: values[1].clone(),
        _prototype: values[2].clone(),
        _partials: values.iter().skip(3).cloned().collect(),
    })
}

fn fixture(count: usize) -> Value {
    let mut values = vec![Value::Undefined, Value::Null, Value::Undefined];
    values.extend((0..count).map(|_| Value::String("captured".into())));
    let mut wrapper = Object::new();
    wrapper.properties.insert("name".into(), Value::String("bound capture".into()));
    wrapper.properties.insert("__bound_args".into(),
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(values))));
    Value::Object(vybe_runtime::heap::alloc(wrapper))
}

fn timed(callback: &Value, iterations: usize, optimized: bool) -> Duration {
    let start = Instant::now();
    for _ in 0..iterations {
        if optimized {
            black_box(vybe_platform_ecma::function::prepare_bound_callback(black_box(callback)));
        } else {
            black_box(previous(black_box(callback)));
        }
    }
    start.elapsed()
}

fn main() {
    let iterations = std::env::args().nth(1)
        .map(|s| s.parse::<usize>().expect("iterations must be a positive integer"))
        .unwrap_or(20_000);
    assert!(iterations > 0);
    for count in [0, 1, 3, 8, 9, 32] {
        let callback = fixture(count);
        assert!(previous(&callback).is_some());
        assert!(vybe_platform_ecma::function::prepare_bound_callback(&callback).is_some());
        let mut old = Duration::ZERO;
        let mut new = Duration::ZERO;
        for round in 0..8 {
            for optimized in if round % 2 == 0 { [false, true] } else { [true, false] } {
                let elapsed = timed(&callback, iterations, optimized);
                if optimized { new += elapsed; } else { old += elapsed; }
            }
        }
        println!("bound decode captures={count}: previous={old:?}, current={new:?}, ratio={:.3}",
            old.as_secs_f64() / new.as_secs_f64());
    }
}
