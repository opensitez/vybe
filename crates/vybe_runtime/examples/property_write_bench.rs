//! Direct property-storage benchmark; no frontend or guest compilation.
use std::hint::black_box;
use std::time::{Duration, Instant};
use vybe_runtime::value::Object;
use vybe_runtime::Value;

fn measure(iterations: usize, borrowed: bool) -> (Duration, f64) {
    let keys: Vec<String> = (0..16).map(|index| format!("field_{index}")).collect();
    let mut object = Object::new();
    for key in &keys {
        object.set(key.clone(), Value::I32(0));
    }
    let start = Instant::now();
    for index in 0..iterations {
        let key = black_box(keys[index % keys.len()].as_str());
        let value = Value::I32(black_box(index as i32));
        if borrowed {
            object.set_borrowed(key, value);
        } else {
            object.set(key.to_owned(), value);
        }
    }
    let elapsed = start.elapsed();
    let checksum = keys.iter().map(|key| object.get(key).as_f64()).sum();
    black_box(&object);
    (elapsed, checksum)
}

fn main() {
    let iterations = std::env::args().nth(1)
        .map(|value| value.parse::<usize>().expect("integer iteration count"))
        .unwrap_or(1_000_000);
    let mut owned = Vec::new();
    let mut borrowed = Vec::new();
    for repeat in 0..7 {
        let (before, after) = if repeat % 2 == 0 {
            (measure(iterations, false), measure(iterations, true))
        } else {
            let after = measure(iterations, true);
            (measure(iterations, false), after)
        };
        assert_eq!(before.1, after.1, "storage results differ");
        owned.push(before.0);
        borrowed.push(after.0);
    }
    owned.sort_unstable();
    borrowed.sort_unstable();
    println!("iterations={iterations} owned_ms={:.3} borrowed_ms={:.3} speedup={:.3}",
        owned[3].as_secs_f64() * 1000.0,
        borrowed[3].as_secs_f64() * 1000.0,
        owned[3].as_secs_f64() / borrowed[3].as_secs_f64());
}
