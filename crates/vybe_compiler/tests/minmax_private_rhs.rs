use vybe_compiler::primitives::{self, collections, ops};
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/minmax_private_before.rs"]
mod before;
#[path = "fixtures/generic_cmp_before.rs"]
mod cmp_before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool, max: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("minmax-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = match (fast, max) {
        (true, false) => collections::build_min(&mut root),
        (true, true) => collections::build_max(&mut root),
        (false, false) => before::build_min(&mut root),
        _ => before::build_max(&mut root),
    };
    c.name = "entry".into();
    c.param_count = c.arity;
    let mut cs = vec![root, c];
    primitives::bundle::finalize_with_runtime_helpers(&mut cs);
    primitives::globals::declare_free_globals(&mut cs);
    primitives::globals::normalize_global_table(&mut cs);
    cs
}
fn worker(fast: bool, max: bool) -> (VM, Value) {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(program(fast, max)).unwrap();
    (vm, f)
}
fn array(v: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(
        vybe_runtime::value::Object::new_array(v),
    ))
}
fn object(ticks: f64) -> Value {
    let mut o = vybe_runtime::value::Object::new();
    o.properties.insert("Ticks".into(), Value::F64(ticks));
    Value::Object(vybe_runtime::heap::alloc(o))
}
fn same(a: &Value, b: &Value) {
    match (a, b) {
        (Value::F64(a), Value::F64(b)) => {
            if a.is_nan() {
                assert!(b.is_nan());
            } else {
                assert_eq!(a.to_bits(), b.to_bits());
            }
        }
        _ => assert_eq!(a, b),
    }
}
#[test]
fn generic_emission_is_unchanged() {
    type E = fn(&mut Chunk, &mut Chunk, u32);
    for (old, new) in [
        (cmp_before::emit_lt as E, ops::emit_dyn_lt_into as E),
        (cmp_before::emit_gt, ops::emit_dyn_gt_into),
        (cmp_before::emit_le, ops::emit_dyn_le_into),
        (cmp_before::emit_ge, ops::emit_dyn_ge_into),
    ] {
        let build = |f: E| {
            let mut imports = Chunk::new("imports");
            let mut c = Chunk::new("compare");
            c.arity = 2;
            c.local_count = 2;
            c.emit_op_u16(Op::LOCAL_GET, 0, 0);
            c.emit_op_u16(Op::LOCAL_GET, 1, 0);
            f(&mut imports, &mut c, 0);
            c.emit_op(Op::RETURN, 0);
            c
        };
        let a = build(old);
        let b = build(new);
        assert_eq!(a.code, b.code);
        assert_eq!(a.local_count, b.local_count);
        assert_eq!(a.constants, b.constants);
        assert_eq!(format!("{:?}", a.imports), format!("{:?}", b.imports));
    }
}
#[test]
fn values_and_sizes() {
    let mut checks = 0;
    let mut rows = vec![];
    for max in [false, true] {
        let mut ws = [worker(false, max), worker(true, max)];
        for len in 0..65 {
            for shape in 0..4 {
                let input: Vec<Value> = (0..len)
                    .map(|i| {
                        Value::I32(match shape {
                            0 => i,
                            1 => len - i,
                            2 => (i * 17 + len) % 11,
                            _ => 3,
                        })
                    })
                    .collect();
                let a = array(input.clone());
                let old_fn = ws[0].1.clone();
                let new_fn = ws[1].1.clone();
                let old = ws[0].0.invoke(&old_fn, &[a.clone()]).unwrap();
                let new = ws[1].0.invoke(&new_fn, &[a]).unwrap();
                same(&old, &new);
                if len > 0 {
                    let expected = if max {
                        input.iter().max_by_key(|v| v.as_f64() as i32)
                    } else {
                        input.iter().min_by_key(|v| v.as_f64() as i32)
                    }
                    .unwrap();
                    same(&old, expected);
                }
                checks += 2;
            }
        }
        for input in [
            vec![Value::F64(f64::NAN), Value::I32(1)],
            vec![Value::I32(1), Value::F64(f64::NAN)],
            vec![Value::F64(-0.0), Value::F64(0.0)],
            vec![Value::String("b".into()), Value::String("a".into())],
            vec![Value::bigint_i64(5), Value::bigint_i64(2)],
            vec![Value::Null, Value::Bool(true), Value::Undefined],
            vec![object(3.0), object(1.0), object(9.0)],
            vec![Value::I32(2), Value::String("1".into())],
            vec![Value::bigint_i64(1), Value::I32(2)],
        ] {
            let a = array(input);
            let mut results = vec![];
            for fast in [false, true] {
                let (mut vm, f) = worker(fast, max);
                results.push(vm.invoke(&f, &[a.clone()]).map_err(|e| e.message.clone()));
            }
            match (&results[0], &results[1]) {
                (Ok(a), Ok(b)) => same(a, b),
                (Err(a), Err(b)) => assert_eq!(a, b),
                _ => panic!("outcome mismatch"),
            };
            checks += 1;
        }
        let mut sizes = vec![];
        for fast in [false, true] {
            let cs = program(fast, max);
            let bytes = vybe_platform_wasm::write_wasm_with_options(
                &cs,
                vybe_platform_wasm::WasmWriteOptions {
                    include_name_section: false,
                    include_vybe_metadata: false,
                },
            );
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&bytes)
                .unwrap();
            if std::env::var_os("VYBE_MINMAX_EXPORT").is_some() {
                native_export::export_chunks(&cs, &format!("minmax-private-{max}-{fast}"));
            }
            sizes.push(serde_json::json!({"body":cs[1].code.len(),"locals":cs[1].local_count,"native":bytes.len()}));
        }
        assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
        assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
        rows.push(serde_json::json!({"max":max,"sizes":sizes}));
    }
    let result = serde_json::json!({"checks":checks,"rows":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../MINMAX_PRIVATE_GATES.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::{Duration, Instant};
    let mut rows = vec![];
    for (max, len, descending, control) in [
        (false, 256, false, true),
        (false, 0, false, false),
        (true, 1, false, false),
        (false, 256, false, false),
        (false, 256, true, false),
        (true, 256, false, false),
        (true, 256, true, false),
    ] {
        let mut vm = VM::new();
        primitives::platforms::register_platforms_all(&mut vm);
        let old = vm.run(program(false, max)).unwrap();
        let new = if control {
            old.clone()
        } else {
            vm.run(program(true, max)).unwrap()
        };
        let fs = [old, new];
        let input = array(
            (0..len)
                .map(|i| Value::I32(if descending { len - i } else { i }))
                .collect(),
        );
        for f in &fs {
            for _ in 0..8 {
                std::hint::black_box(vm.invoke(f, &[input.clone()]).unwrap());
            }
        }
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let t = Instant::now();
                let mut n = 0;
                while t.elapsed() < Duration::from_millis(500) || n < 8 {
                    std::hint::black_box(vm.invoke(&fs[side], &[input.clone()]).unwrap());
                    n += 1;
                }
                samples[side].push(t.elapsed().as_secs_f64() / n as f64);
            }
        }
        let mut ratios: Vec<_> = samples[0]
            .iter()
            .zip(&samples[1])
            .map(|(a, b)| a / b)
            .collect();
        ratios.sort_by(f64::total_cmp);
        rows.push(serde_json::json!({"max":max,"len":len,"descending":descending,"control":control,"median_speedup":ratios[3],"seconds_per_call":samples}));
    }
    let result = serde_json::json!({"same_vm":true,"pairs":7,"minimum_sample_ms":500,"minimum_calls":8,"warmups":8,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../MINMAX_PRIVATE_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}

#[cfg(feature = "primitive-benchmarks")]
fn natural_program(fast: bool, copy: bool) -> Vec<Chunk> {
    ops::benchmark_with_comparison_slots(fast, || {
        let mut root = Chunk::new("natural-owner");
        primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
        root.emit_op(Op::RETURN, 0);
        let mut c = if copy {
            collections::build_sorted(&mut root)
        } else {
            collections::build_sort_in_place(&mut root)
        };
        c.name = "entry".into();
        c.param_count = c.arity;
        let mut cs = vec![root, c];
        primitives::bundle::finalize_with_runtime_helpers(&mut cs);
        primitives::globals::declare_free_globals(&mut cs);
        primitives::globals::normalize_global_table(&mut cs);
        cs
    })
}
#[cfg(feature = "primitive-benchmarks")]
#[test]
fn natural_sort_values_and_sizes() {
    let mut checks = 0;
    let mut rows = vec![];
    for copy in [false, true] {
        let mut ws: Vec<_> = [false, true]
            .into_iter()
            .map(|fast| {
                let mut vm = VM::new();
                primitives::platforms::register_platforms_all(&mut vm);
                let f = vm.run(natural_program(fast, copy)).unwrap();
                (vm, f)
            })
            .collect();
        for len in 0..33 {
            for shape in 0..3 {
                let input: Vec<Value> = (0..len)
                    .map(|i| {
                        Value::I32(match shape {
                            0 => i,
                            1 => len - i,
                            _ => (i * 17 + len) % 7,
                        })
                    })
                    .collect();
                let mut expected = input.clone();
                expected.sort_by_key(|v| v.as_f64() as i32);
                for (vm, f) in &mut ws {
                    let a = array(input.clone());
                    let out = vm.invoke(f, &[a.clone()]).unwrap();
                    let read = |v: &Value| {
                        let Value::Object(o) = v else { panic!() };
                        let o = o.lock().unwrap();
                        let vybe_runtime::value::ObjectKind::Array(v) = &o.kind else {
                            panic!()
                        };
                        v.clone()
                    };
                    assert_eq!(read(&out), expected);
                    assert_eq!(
                        read(&a),
                        if copy {
                            input.clone()
                        } else {
                            expected.clone()
                        }
                    );
                    checks += 2;
                }
            }
        }
        let mut sizes = vec![];
        for fast in [false, true] {
            let cs = natural_program(fast, copy);
            let bytes = vybe_platform_wasm::write_wasm_with_options(
                &cs,
                vybe_platform_wasm::WasmWriteOptions {
                    include_name_section: false,
                    include_vybe_metadata: false,
                },
            );
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&bytes)
                .unwrap();
            if std::env::var_os("VYBE_MINMAX_EXPORT").is_some() {
                native_export::export_chunks(&cs, &format!("natural-private-{copy}-{fast}"));
            }
            sizes.push(serde_json::json!({"body":cs[1].code.len(),"locals":cs[1].local_count,"native":bytes.len()}));
        }
        assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
        assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
        rows.push(serde_json::json!({"copy":copy,"sizes":sizes}));
    }
    let result = serde_json::json!({"checks":checks,"rows":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../NATURAL_PRIVATE_GATES.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
#[cfg(feature = "primitive-benchmarks")]
#[test]
#[ignore]
fn debug_sort_pairs() {
    use std::time::{Duration, Instant};
    let mut rows = vec![];
    for (copy, len, shape, control) in [
        (false, 256, 2, true),
        (false, 256, 0, false),
        (false, 256, 2, false),
        (true, 256, 1, false),
    ] {
        let mut vm = VM::new();
        primitives::platforms::register_platforms_all(&mut vm);
        let old = vm.run(natural_program(false, copy)).unwrap();
        let new = if control {
            old.clone()
        } else {
            vm.run(natural_program(true, copy)).unwrap()
        };
        let fs = [old, new];
        let input: Vec<Value> = (0..len)
            .map(|i| {
                Value::I32(match shape {
                    0 => i,
                    1 => len - i,
                    _ => (i * 17 + 3) % 101,
                })
            })
            .collect();
        for f in &fs {
            for _ in 0..8 {
                std::hint::black_box(vm.invoke(f, &[array(input.clone())]).unwrap());
            }
        }
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let t = Instant::now();
                let mut n = 0;
                while t.elapsed() < Duration::from_millis(500) || n < 8 {
                    std::hint::black_box(vm.invoke(&fs[side], &[array(input.clone())]).unwrap());
                    n += 1;
                }
                samples[side].push(t.elapsed().as_secs_f64() / n as f64);
            }
        }
        let mut ratios: Vec<_> = samples[0]
            .iter()
            .zip(&samples[1])
            .map(|(a, b)| a / b)
            .collect();
        ratios.sort_by(f64::total_cmp);
        rows.push(serde_json::json!({"copy":copy,"len":len,"shape":shape,"control":control,"median_speedup":ratios[3],"seconds_per_call":samples}));
    }
    let result = serde_json::json!({"same_vm":true,"pairs":7,"minimum_sample_ms":500,"minimum_calls":8,"warmups":8,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../NATURAL_PRIVATE_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
