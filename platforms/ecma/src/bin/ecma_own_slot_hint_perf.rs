//! Compiler-free experiment for a validated own-property index hint.
//! No production cache or type/descriptor machinery is changed. The hint
//! contains no Values or object ownership and rechecks the current key on hits.
//! Timings include thread-local hint access, but not complete ECMA Get semantics.

use std::{cell::Cell, hint::black_box, time::{Duration, Instant}};
use vybe_runtime::{value::Object, Value};

#[derive(Clone, Copy, Default)]
struct SlotHint { owner: usize, index: Option<usize> }

thread_local! { static HINT: Cell<SlotHint> = Cell::new(SlotHint::default()); }
thread_local! { static TWO_HINTS: Cell<[Option<usize>; 2]> = const { Cell::new([None; 2]) }; }

fn two_hinted<'a>(object: &'a Object, key: &str) -> Option<&'a Value> {
    if object.properties.len() <= 1 {
        return object.properties.get(key);
    }
    TWO_HINTS.with(|cache| {
        let hints = cache.get();
        for index in hints.into_iter().flatten() {
            if let Some((stored, value)) = object.properties.get_index(index) {
                if stored == key { return Some(value); }
            }
        }
        let result = object.properties.get_full(key);
        if let Some((index, _, _)) = result {
            cache.set([Some(index), hints[0]]);
        }
        result.map(|(_, _, value)| value)
    })
}

fn hinted<'a>(object: &'a Object, key: &str) -> Option<&'a Value> {
    HINT.with(|cache| {
        let owner = object as *const Object as usize;
        let hint = cache.get();
        if hint.owner == owner {
            if let Some(index) = hint.index {
                if let Some((stored, value)) = object.properties.get_index(index) {
                    if stored == key { return Some(value); }
                }
            }
        }
        let result = object.properties.get_full(key);
        cache.set(SlotHint { owner, index: result.map(|(index, _, _)| index) });
        result.map(|(_, _, value)| value)
    })
}

fn fixture(count: usize) -> Object {
    let mut object = Object::new();
    for n in 0..count {
        object.properties.insert(format!("applicationField{n}"), Value::I32(n as i32));
    }
    object
}

fn verify_mutation_guards() {
    let mut object = fixture(3);
    assert!(matches!(hinted(&object, "applicationField1"), Some(Value::I32(1))));
    object.properties.shift_remove("applicationField0");
    assert!(matches!(hinted(&object, "applicationField1"), Some(Value::I32(1))));
    object.properties.insert("applicationField1".into(), Value::I32(91));
    assert!(matches!(hinted(&object, "applicationField1"), Some(Value::I32(91))));
    object.properties.shift_remove("applicationField1");
    assert!(hinted(&object, "applicationField1").is_none());
    object.properties.insert("applicationField1".into(), Value::I32(92));
    assert!(matches!(hinted(&object, "applicationField1"), Some(Value::I32(92))));
    object.properties.clear();
    object.properties.insert("other".into(), Value::I32(93));
    assert!(hinted(&object, "applicationField1").is_none());
    assert!(matches!(hinted(&object, "other"), Some(Value::I32(93))));
    let mut first = fixture(3);
    let second = fixture(3);
    assert!(matches!(two_hinted(&first, "applicationField1"), Some(Value::I32(1))));
    assert!(matches!(two_hinted(&second, "applicationField1"), Some(Value::I32(1))));
    first.properties.shift_remove("applicationField0");
    assert!(matches!(two_hinted(&first, "applicationField1"), Some(Value::I32(1))));
    first.properties.insert("applicationField1".into(), Value::I32(99));
    assert!(matches!(two_hinted(&first, "applicationField1"), Some(Value::I32(99))));
    first.properties.shift_remove("applicationField1");
    assert!(two_hinted(&first, "applicationField1").is_none());
    first.properties.clear();
    first.properties.insert("other".into(), Value::I32(98));
    assert!(two_hinted(&first, "applicationField1").is_none());
    assert!(matches!(two_hinted(&first, "other"), Some(Value::I32(98))));
}

fn timed(objects: &[Object], keys: &[String], iterations: usize, variant: usize) -> Duration {
    HINT.with(|cache| cache.set(SlotHint::default()));
    TWO_HINTS.with(|cache| cache.set([None; 2]));
    let start = Instant::now();
    for iteration in 0..iterations {
        let object = &objects[iteration % objects.len()];
        for key in keys {
            let object = black_box(object);
            let key = black_box(key.as_str());
            black_box(match variant {
                1 => hinted(object, key),
                2 => two_hinted(object, key),
                _ => object.properties.get(key),
            });
        }
    }
    start.elapsed()
}

fn main() {
    verify_mutation_guards();
    let iterations = std::env::args().nth(1)
        .map(|s| s.parse::<usize>().expect("iterations must be a positive integer"))
        .unwrap_or(50_000);
    assert!(iterations > 0);
    for count in [1, 3, 16, 64, 512] {
        let objects = [fixture(count), fixture(count), fixture(count)];
        for (mode, keys) in [
            ("repeated", vec![format!("applicationField{}", count - 1)]),
            ("alternating", vec!["applicationField0".into(), format!("applicationField{}", count - 1)]),
            ("three_keys", (0..count.min(3)).map(|n| format!("applicationField{n}")).collect()),
            ("miss", vec!["absentField".into()]),
        ] {
            for cross_objects in [false, true] {
                let input = if cross_objects { &objects[..] } else { &objects[..1] };
                let mut totals = [Duration::ZERO; 3];
                for round in 0..8 {
                    for variant in if round % 2 == 0 { [0, 1, 2] } else { [2, 1, 0] } {
                        totals[variant] += timed(input, &keys, iterations, variant);
                    }
                }
                println!("own slot count={count} mode={mode} cross_objects={cross_objects}: ordinary={:?}, single={:?}, two={:?}, single_ratio={:.3}, two_ratio={:.3}",
                    totals[0], totals[1], totals[2], totals[0].as_secs_f64() / totals[1].as_secs_f64(),
                    totals[0].as_secs_f64() / totals[2].as_secs_f64());
            }
        }
    }
}
