use vybe_compiler::primitives::{self, collections};
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/sum_accumulate_before.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("sum-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = if fast {
        collections::build_sum(&mut root)
    } else {
        before::build_sum(&mut root)
    };
    c.name = "entry".into();
    c.param_count = c.arity;
    let mut cs = vec![root, c];
    primitives::bundle::finalize_with_runtime_helpers(&mut cs);
    primitives::globals::declare_free_globals(&mut cs);
    primitives::globals::normalize_global_table(&mut cs);
    cs
}
fn worker(fast: bool) -> (VM, Value) {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(program(fast)).unwrap();
    (vm, f)
}
fn array(v: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(
        vybe_runtime::value::Object::new_array(v),
    ))
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
fn invoke(w: &mut (VM, Value), v: Value) -> Value {
    w.0.invoke(&w.1, &[v]).unwrap()
}
#[test]
fn totals_sequences_and_sizes() {
    let mut checks = 0;
    let mut ws = [worker(false), worker(true)];
    for len in 0..130 {
        let nums: Vec<Value> = (0..len)
            .map(|i| Value::F64((i as f64 % 17.0) - 8.25))
            .collect();
        let expected = nums.iter().fold(0.0, |a, b| a + b.as_f64());
        for w in &mut ws {
            assert_eq!(invoke(w, array(nums.clone())).as_f64(), expected);
            checks += 1;
        }
    }
    let alphabet = [
        Value::I32(-2),
        Value::F64(0.25),
        Value::String("".into()),
        Value::String("x".into()),
        Value::Bool(true),
        Value::Null,
        Value::Undefined,
    ];
    for len in 0..5 {
        for mut bits in 0..7usize.pow(len) {
            let mut v = vec![];
            for _ in 0..len {
                v.push(alphabet[bits % 7].clone());
                bits /= 7;
            }
            let a = invoke(&mut ws[0], array(v.clone()));
            let b = invoke(&mut ws[1], array(v));
            same(&a, &b);
            checks += 1;
        }
    }
    for v in [
        vec![],
        vec![Value::F64(-0.0)],
        vec![Value::F64(f64::NAN), Value::String("x".into())],
        vec![Value::F64(f64::INFINITY), Value::String("x".into())],
        vec![Value::Bool(false), Value::String("é😀".into()), Value::Null],
    ] {
        let a = invoke(&mut ws[0], array(v.clone()));
        let b = invoke(&mut ws[1], array(v));
        same(&a, &b);
        checks += 1;
    }
    for prefix in [
        vec![],
        vec![Value::I32(3)],
        vec![Value::String("x".into())],
        vec![Value::Null, Value::Undefined],
    ] {
        for bad in [
            Value::bigint_i64(7),
            Value::Object(vybe_runtime::heap::alloc(vybe_runtime::value::Object::new())),
        ] {
            let mut input = prefix.clone();
            input.push(bad);
            input.push(Value::String("after".into()));
            let mut results = vec![];
            for fast in [false, true] {
                let (mut vm, f) = worker(fast);
                results.push(
                    vm.invoke(&f, &[array(input.clone())])
                        .map_err(|e| e.message.clone()),
                );
            }
            assert!(results[0].is_err());
            assert_eq!(results[0], results[1]);
            checks += 1;
        }
    }
    let mut sizes = vec![];
    for fast in [false, true] {
        let cs = program(fast);
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
        if std::env::var_os("VYBE_SUM_EXPORT").is_some() {
            native_export::export_chunks(&cs, &format!("sum-accumulate-{fast}"));
        }
        sizes.push(serde_json::json!({"body":cs[1].code.len(),"bodies":cs.iter().map(|c|c.code.len()).sum::<usize>(),"locals":cs[1].local_count,"native":bytes.len()}));
    }
    assert!(sizes[1]["bodies"].as_u64() < sizes[0]["bodies"].as_u64());
    assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
    let result = serde_json::json!({"checks":checks,"sizes":sizes});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../SUM_ACCUMULATE_GATES.json"),
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
    for (len, strings, control) in [
        (256, false, true),
        (0, false, false),
        (8, false, false),
        (256, false, false),
        (4096, false, false),
        (64, true, false),
    ] {
        let mut vm = VM::new();
        primitives::platforms::register_platforms_all(&mut vm);
        let old = vm.run(program(false)).unwrap();
        let new = if control {
            old.clone()
        } else {
            vm.run(program(true)).unwrap()
        };
        let fs = [old, new];
        let input = array(
            (0..len)
                .map(|i| {
                    if strings && i % 3 == 0 {
                        Value::String("x".into())
                    } else {
                        Value::I32(i % 17)
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
        rows.push(serde_json::json!({"len":len,"strings":strings,"control":control,"median_speedup":ratios[3],"seconds_per_call":samples}));
    }
    let result = serde_json::json!({"same_vm":true,"pairs":7,"minimum_sample_ms":500,"minimum_calls":8,"warmups":8,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../SUM_ACCUMULATE_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
