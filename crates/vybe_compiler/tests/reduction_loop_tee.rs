#![cfg(feature = "primitive-benchmarks")]
use vybe_compiler::primitives::{self, collections, instructions::core_wasm};
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/any_all_loop_before.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool, kind: usize) -> Vec<Chunk> {
    core_wasm::benchmark_with_local_tee(fast, || {
        let mut root = Chunk::new("reduction-owner");
        primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
        root.emit_op(Op::RETURN, 0);
        let mut c = if kind == 2 {
            collections::build_sum(&mut root)
        } else if fast {
            collections::build_any_all(
                &mut root,
                if kind == 0 {
                    "__stdlib_pyany"
                } else {
                    "__stdlib_pyall"
                },
                kind == 0,
            )
        } else {
            before::build_any_all(
                &mut root,
                if kind == 0 {
                    "__stdlib_pyany"
                } else {
                    "__stdlib_pyall"
                },
                kind == 0,
            )
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
fn worker(fast: bool, kind: usize) -> (VM, Value) {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(program(fast, kind)).unwrap();
    (vm, f)
}
fn array(v: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(
        vybe_runtime::value::Object::new_array(v),
    ))
}
#[test]
fn truthiness_totals_and_sizes() {
    let mut checks = 0;
    let mut rows = vec![];
    let alphabet = [
        (Value::I32(0), false),
        (Value::I32(1), true),
        (Value::F64(f64::NAN), false),
        (Value::String("".into()), false),
        (Value::String("x".into()), true),
        (Value::Bool(false), false),
        (Value::Bool(true), true),
        (Value::Null, false),
        (Value::Undefined, false),
        (Value::bigint_i64(0), false),
        (Value::bigint_i64(1), true),
        (
            Value::Object(vybe_runtime::heap::alloc(vybe_runtime::value::Object::new())),
            true,
        ),
    ];
    for kind in 0..3 {
        let mut ws = [worker(false, kind), worker(true, kind)];
        if kind < 2 {
            for len in 0..4 {
                for mut code in 0..12usize.pow(len) {
                    let mut a = vec![];
                    let mut truth = vec![];
                    for _ in 0..len {
                        let (v, t) = &alphabet[code % 12];
                        a.push(v.clone());
                        truth.push(*t);
                        code /= 12;
                    }
                    let expected = if kind == 0 {
                        truth.iter().any(|v| *v)
                    } else {
                        truth.iter().all(|v| *v)
                    };
                    for (vm, f) in &mut ws {
                        assert_eq!(
                            vm.invoke(f, &[array(a.clone())]).unwrap(),
                            Value::Bool(expected)
                        );
                        checks += 1;
                    }
                }
            }
        } else {
            for len in 0..130 {
                let a: Vec<Value> = (0..len).map(Value::I32).collect();
                for (vm, f) in &mut ws {
                    assert_eq!(
                        vm.invoke(f, &[array(a.clone())]).unwrap().as_f64(),
                        (len * (len - 1) / 2) as f64
                    );
                    checks += 1;
                }
            }
            for (vm, f) in &mut ws {
                assert_eq!(
                    vm.invoke(
                        f,
                        &[array(vec![
                            Value::String("x".into()),
                            Value::I32(2),
                            Value::Null
                        ])]
                    )
                    .unwrap(),
                    Value::String("0x20".into())
                );
                checks += 1;
            }
        }
        let mut sizes = vec![];
        for fast in [false, true] {
            let cs = program(fast, kind);
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
            if std::env::var_os("VYBE_REDUCTION_TEE_EXPORT").is_some() {
                native_export::export_chunks(&cs, &format!("reduction-tee-{kind}-{fast}"));
            }
            sizes.push(serde_json::json!({"body":cs[1].code.len(),"locals":cs[1].local_count,"native":bytes.len()}));
        }
        assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
        eprintln!("kind={kind} sizes={sizes:?}");
        if kind == 2 { assert!(sizes[1]["native"].as_u64() <= sizes[0]["native"].as_u64()); }
        rows.push(serde_json::json!({"kind":kind,"sizes":sizes}));
    }
    let result = serde_json::json!({"checks":checks,"rows":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../REDUCTION_TEE_GATES.json"),
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
    for (kind, shape, control) in [
        (0, 0, true),
        (0, 1, false),
        (0, 2, false),
        (0, 3, false),
        (1, 0, false),
        (1, 2, false),
        (2, 0, false),
        (2, 4, false),
    ] {
        let (mut vm, _) = worker(false, kind);
        let a = vm.run(program(false, kind)).unwrap();
        let b = if control {
            a.clone()
        } else {
            vm.run(program(true, kind)).unwrap()
        };
        let fs = [a, b];
        let len = if shape == 1 {
            0
        } else if shape == 4 {
            64
        } else {
            256
        };
        let input = array(
            (0..len)
                .map(|i| {
                    if kind == 2 {
                        if shape == 4 && i % 3 == 0 {
                            Value::String("x".into())
                        } else {
                            Value::I32(i % 17)
                        }
                    } else {
                        Value::Bool(if kind == 0 {
                            (shape == 2 && i == 0) || (shape == 3 && i == len - 1)
                        } else {
                            !(shape == 2 && i == 0)
                        })
                    }
                })
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
        rows.push(serde_json::json!({"kind":kind,"shape":shape,"control":control,"len":len,"median_speedup":ratios[3],"seconds_per_call":samples}));
    }
    let result = serde_json::json!({"same_vm":true,"pairs":7,"minimum_sample_ms":500,"minimum_calls":8,"warmups":8,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../REDUCTION_TEE_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
