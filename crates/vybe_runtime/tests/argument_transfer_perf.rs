//! Paired argument-tail transfer helper timing, not complete VM or Doom timing.
use std::hint::black_box;
use std::sync::{Arc, Mutex};
use std::time::Instant;
use vybe_runtime::value::Object;
use vybe_runtime::Value;

fn transfer(stack: &mut Vec<Value>, moving: bool) -> Vec<Value> {
    if moving {
        stack.split_off(1)
    } else {
        let args = stack[1..].to_vec();
        stack.truncate(1);
        args
    }
}

fn time(seed: &[Value], moving: bool, rounds: usize) -> f64 {
    let mut stack = Vec::with_capacity(seed.len() + 8);
    stack.push(Value::Undefined);
    stack.extend_from_slice(seed);
    let start = Instant::now();
    for _ in 0..rounds {
        let args = transfer(black_box(&mut stack), black_box(moving));
        black_box(&args);
        // Restore by ownership transfer, identically for both variants.
        stack.extend(args);
    }
    let elapsed = start.elapsed().as_secs_f64();
    assert_eq!(stack.len(), seed.len() + 1);
    elapsed
}

fn cases() -> Vec<(&'static str, Value)> {
    vec![
        ("i32", Value::I32(42)),
        ("v128", Value::V128([0x5a; 16])),
        ("string", Value::String(Arc::from("argument-transfer-reference"))),
        ("object", Value::Object(Arc::new(Mutex::new(Object::new_array(vec![Value::I32(42)]))))),
    ]
}

#[test]
fn transferred_arguments_preserve_values_identity_and_release() {
    for (_, value) in cases() {
        for moving in [false, true] {
            let mut stack = vec![Value::Undefined, value.clone(), value.clone()];
            let args = transfer(&mut stack, moving);
            assert_eq!(stack, vec![Value::Undefined]);
            assert_eq!(args, vec![value.clone(), value.clone()]);
            match &value {
                Value::Object(expected) => {
                    for arg in &args {
                        let Value::Object(actual) = arg else { panic!("object representation lost"); };
                        assert!(Arc::ptr_eq(expected, actual));
                    }
                    assert_eq!(Arc::strong_count(expected), 3);
                    drop(args);
                    assert_eq!(Arc::strong_count(expected), 1);
                }
                Value::String(expected) => {
                    for arg in &args {
                        let Value::String(actual) = arg else { panic!("string representation lost"); };
                        assert!(Arc::ptr_eq(expected, actual));
                    }
                    assert_eq!(Arc::strong_count(expected), 3);
                    drop(args);
                    assert_eq!(Arc::strong_count(expected), 1);
                }
                _ => {}
            }
        }
    }
}

#[test]
#[ignore = "paired release helper benchmark; not a VM or Doom speedup"]
fn paired_argument_transfer() {
    for (kind, value) in cases() {
        for count in [0, 1, 4, 16, 64] {
            let seed = vec![value.clone(); count];
            time(&seed, false, 1000);
            time(&seed, true, 1000);
            let mut ratios = Vec::with_capacity(11);
            for pair in 0..11 {
                let (old, new) = if pair % 2 == 0 {
                    (time(&seed, false, 20_000), time(&seed, true, 20_000))
                } else {
                    let new = time(&seed, true, 20_000);
                    (time(&seed, false, 20_000), new)
                };
                ratios.push(old / new);
            }
            ratios.sort_by(f64::total_cmp);
            eprintln!("argument-transfer kind={kind} count={count} old/new median={:.4} pairs={ratios:?}", ratios[5]);
        }
    }
}

fn time_inline(seed: &[Value], rounds: usize) -> f64 {
    assert!(seed.len() <= 4);
    let mut stack = Vec::with_capacity(seed.len() + 8);
    stack.push(Value::Undefined);
    stack.extend_from_slice(seed);
    let start = Instant::now();
    for _ in 0..rounds {
        let mut args = [const { Value::Undefined }; 4];
        for argument in args[..seed.len()].iter_mut().rev() {
            *argument = black_box(&mut stack).pop().unwrap();
        }
        black_box(&args[..seed.len()]);
        stack.extend(args.into_iter().take(seed.len()));
    }
    let elapsed = start.elapsed().as_secs_f64();
    assert_eq!(stack.len(), seed.len() + 1);
    elapsed
}

#[test]
#[ignore = "paired inline/Vec release helper benchmark; not complete VM timing"]
fn paired_inline_argument_transfer() {
    for (kind, value) in cases() {
        for count in [0, 1, 2, 4] {
            let seed = vec![value.clone(); count];
            time(&seed, true, 1000);
            time_inline(&seed, 1000);
            let mut ratios = Vec::with_capacity(11);
            for pair in 0..11 {
                let (vec_time, inline_time) = if pair % 2 == 0 {
                    (time(&seed, true, 200_000), time_inline(&seed, 200_000))
                } else {
                    let inline_time = time_inline(&seed, 200_000);
                    (time(&seed, true, 200_000), inline_time)
                };
                ratios.push(vec_time / inline_time);
            }
            ratios.sort_by(f64::total_cmp);
            eprintln!("inline-argument-transfer kind={kind} count={count} Vec/inline median={:.4} pairs={ratios:?}", ratios[5]);
        }
    }
}
