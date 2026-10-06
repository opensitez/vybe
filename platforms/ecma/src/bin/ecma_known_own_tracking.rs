//! Matched overwrite-loop benchmark, not complete Object.assign timing.
use std::hint::black_box;
use std::time::Instant;
use vybe_runtime::value::{Object, Value};

fn overwrite(target: &mut Object, keys: &[String], values: &[Value], legacy: bool) {
    for (index, (key, value)) in keys.iter().zip(values).enumerate() {
        if legacy {
            // This is the real helper used by the previous overwrite batch.
            // All keys already exist; no insertion bookkeeping is required.
            let tracking = vybe_platform_ecma::object::prepare_track_key_in(target, key);
            assert!(tracking.is_none());
        }
        let slot = target.properties.get_index_mut(index)
            .filter(|(stored, _)| stored.as_str() == key.as_str());
        if let Some((_, slot)) = slot {
            *slot = value.clone();
        } else {
            *target.properties.get_mut(key).expect("existing own key") = value.clone();
        }
    }
}

fn main() {
    for count in [16usize, 64, 256, 1024] {
        let keys: Vec<String> = (0..count).map(|index| format!("field{index}")).collect();
        let values: Vec<Value> = (0..count).map(|index| Value::I32(index as i32)).collect();
        let mut target = Object::new();
        for key in &keys {
            target.properties.insert(key.clone(), Value::Undefined);
        }
        for legacy in [true, false] {
            overwrite(&mut target, &keys, &values, legacy);
            for (key, value) in keys.iter().zip(&values) {
                assert!(matches!((target.properties.get(key), value),
                    (Some(Value::I32(actual)), Value::I32(expected)) if actual == expected));
            }
        }
        let iterations = (1_000_000 / count).max(256);
        let mut totals = [0u128; 2];
        for sample in 0..6 {
            let order = if sample % 2 == 0 { [true, false] } else { [false, true] };
            for legacy in order {
                let start = Instant::now();
                for _ in 0..iterations {
                    overwrite(black_box(&mut target), black_box(&keys), black_box(&values), legacy);
                }
                black_box(&target);
                let elapsed = start.elapsed().as_nanos();
                totals[usize::from(!legacy)] += elapsed;
                println!("properties={count} sample={sample} legacy={legacy} ns={elapsed}");
            }
        }
        println!("properties={count} iterations_per_sample={iterations} legacy_over_current={:.3}",
            totals[0] as f64 / totals[1] as f64);
        assert_eq!(target.properties.len(), count);
        assert!(target.properties.keys().eq(keys.iter()));
    }
}
