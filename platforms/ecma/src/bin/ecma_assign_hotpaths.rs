//! Registered Object.assign scaling probe; timings include VM dispatch.
use std::sync::{Arc, Mutex};
use std::time::Instant;
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Value, VM};

fn ordinary() -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new())))
}

// Inner-kernel experiment only: these fixtures contain ordinary data entries,
// not accessors, proxies, symbols, readonly fields, or descriptor metadata.
// Production integration would additionally require the existing assign guards.
fn aligned_overwrite(target: &mut Object, source: &Object) -> bool {
    if target.properties.len() < source.properties.len()
        || !source.properties.keys().zip(target.properties.keys()).all(|(source, target)| source == target)
    {
        return false;
    }
    for ((_, source), (_, target)) in source.properties.iter().zip(target.properties.iter_mut()) {
        *target = source.clone();
    }
    true
}

fn hashed_overwrite(target: &mut Object, source: &Object) {
    for (key, value) in &source.properties {
        *target.properties.get_mut(key).expect("preexisting fixture key") = value.clone();
    }
}

fn probe_overwrite_kernels() {
    for count in [128usize, 1024, 4096] {
        let mut source = Object::new();
        for index in 0..count {
            source.properties.insert(format!("field{index}"), Value::I32(index as i32));
        }
        for reversed in [false, true] {
            let mut target = Object::new();
            for step in 0..count {
                let index = if reversed { count - step - 1 } else { step };
                target.properties.insert(format!("field{index}"), Value::I32(-1));
            }
            // A rejected layout must not have performed partial writes.
            assert_eq!(aligned_overwrite(&mut target, &source), !reversed);
            if reversed {
                assert!(target.properties.values().all(|value| matches!(value, Value::I32(-1))));
                hashed_overwrite(&mut target, &source);
            }
            let iterations = 200_000 / count;
            for sample in 0..4 {
                for optimized in if sample % 2 == 0 { [false, true] } else { [true, false] } {
                    let start = Instant::now();
                    for _ in 0..iterations {
                        if !optimized || !aligned_overwrite(std::hint::black_box(&mut target), std::hint::black_box(&source)) {
                            hashed_overwrite(&mut target, &source);
                        }
                        std::hint::black_box(&target);
                    }
                    println!("overwrite_kernel properties={count} reversed={reversed} sample={sample} optimized={optimized} elapsed_ms={:.3}", start.elapsed().as_secs_f64() * 1000.0);
                }
            }
            for (key, expected) in &source.properties {
                assert_eq!(target.properties.get(key), Some(expected));
            }
            source.properties.insert("field0".into(), Value::I32(-17));
            if !aligned_overwrite(&mut target, &source) {
                hashed_overwrite(&mut target, &source);
            }
            assert!(matches!(target.properties.get("field0"), Some(Value::I32(-17))));
            source.properties.insert("field0".into(), Value::I32(0));
        }
        let mut short = Object::new();
        short.properties.insert("field0".into(), Value::I32(-1));
        assert!(!aligned_overwrite(&mut short, &source));
        assert!(matches!(short.properties.get("field0"), Some(Value::I32(-1))));
    }
}

fn main() {
    let profile_assign = std::env::args().any(|arg| arg == "--profile-assign");
    if !profile_assign {
        probe_overwrite_kernels();
    }
    let mut vm = VM::new();
    vybe_platform_ecma::object::register(&mut vm);
    let index = vm.resolve_host_function_index("ecma:object", "assign")
        .expect("registered Object.assign");
    let mut callable = Object::new();
    callable.kind = ObjectKind::HostFunction(index);
    let callable = Value::Object(Arc::new(Mutex::new(callable)));

    for count in [0usize, 1, 4, 16, 64, 256, 1024].into_iter().filter(|count| !profile_assign || *count == 1024) {
        let source = ordinary();
        let Value::Object(source_object) = &source else { unreachable!() };
        {
            let mut source_object = source_object.lock().unwrap();
            for index in 0..count {
                source_object.properties.insert(format!("field{index}"), Value::I32(index as i32));
            }
        }
        let target = ordinary();
        let arguments = [target.clone(), source.clone()];
        let result = vm.invoke(&callable, &arguments).expect("assign invocation");
        let (Value::Object(result), Value::Object(target_object)) = (&result, &target) else {
            panic!("assign must return the target object");
        };
        assert!(Arc::ptr_eq(result, target_object));
        {
            let target_object = target_object.lock().unwrap();
            for index in 0..count {
                assert!(matches!(target_object.properties.get(&format!("field{index}")), Some(Value::I32(value)) if *value == index as i32));
            }
        }
        if count == 16 || count == 1024 {
            let target_object = target_object.lock().unwrap();
            println!("target_layout properties={count} first_keys={:?}", target_object.properties.keys().take(3).collect::<Vec<_>>());
        }
        // Reuse existing keys to distinguish overwrite cost from growth/setup.
        let iterations = if profile_assign { 8_000_000 } else { 200_000 } / count.max(1);
        let start = Instant::now();
        for _ in 0..iterations {
            std::hint::black_box(vm.invoke(&callable, &arguments).expect("assign overwrite"));
        }
        let elapsed = start.elapsed();
        println!("properties={count} iterations={iterations} elapsed_ms={:.3} us_per_call={:.3} ns_per_property={:.1}",
            elapsed.as_secs_f64() * 1000.0,
            elapsed.as_secs_f64() * 1_000_000.0 / iterations as f64,
            elapsed.as_secs_f64() * 1_000_000_000.0 / (iterations * count.max(1)) as f64);

        // Each invocation must read current source values, not a cached snapshot.
        if count != 0 {
            source_object.lock().unwrap().properties.insert("field0".into(), Value::I32(-17));
            vm.invoke(&callable, &arguments).expect("assign after source mutation");
            assert!(matches!(target_object.lock().unwrap().properties.get("field0"), Some(Value::I32(-17))));
        }
    }
}
