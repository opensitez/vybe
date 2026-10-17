#![cfg(feature = "primitive-benchmarks")]
use std::sync::{Arc, Mutex};
use vybe_compiler::primitives::{self, collections, ops};
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool, kind: usize, repeats: usize) -> Vec<Chunk> {
    ops::benchmark_with_zero_comparison(fast, || {
        let receiver = kind >= 3;
        let mut root = Chunk::new("zero-owner");
        if receiver {
            root.module_receiver_abi = vybe_runtime::chunk::ReceiverAbi::Parameter;
        }
        primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
        root.emit_op(Op::RETURN, 0);
        let mut c = Chunk::new("entry");
        c.arity = 1;
        c.param_count = 1;
        c.local_count = 1;
        let mut cb = Chunk::new("comparator");
        cb.arity = if receiver { 3 } else { 2 };
        cb.param_count = cb.arity;
        cb.local_count = cb.arity as u16;
        cb.emit_op_u16(Op::LOCAL_GET, if receiver { 1 } else { 0 }, 0);
        cb.emit_op_u16(Op::LOCAL_GET, if receiver { 2 } else { 1 }, 0);
        let compare = cb.add_import("test:sort", "compare");
        cb.emit_call(compare, 2, 0);
        cb.emit_op(Op::RETURN, 0);
        let mut cs = if kind == 2 || kind == 4 {
            let mut body = collections::build_sort_with_comparator(&mut root);
            body.param_count = body.arity;
            primitives::functions::emit_ref_func(&mut c, 3, 0, 0);
            c.emit_op_u16(Op::LOCAL_GET, 0, 0);
            primitives::functions::emit_ref_func(&mut c, 2, 0, 0);
            primitives::callable::emit_direct_invoke_chunk(&mut c, 2, 0);
            vec![root, c, cb, body]
        } else {
            vec![root, c, cb]
        };
        if kind == 0 {
            for i in 0..repeats {
                cs[1].emit_op_u16(Op::LOCAL_GET, 0, 0);
                let mut imports = Chunk::new("imports");
                ops::emit_dyn_gt_zero_into(&mut imports, &mut cs[1], 0);
                if i + 1 < repeats {
                    cs[1].emit_op(Op::DROP, 0);
                }
            }
        } else if kind == 1 || kind == 3 {
            cs[1].emit_op_u16(Op::LOCAL_GET, 0, 0);
            primitives::functions::emit_ref_func(&mut cs[1], 2, 0, 0);
            collections::emit_sort_func(&mut cs, 1, 0);
        }
        cs[1].emit_op(Op::RETURN, 0);
        primitives::bundle::finalize_with_runtime_helpers(&mut cs);
        primitives::globals::declare_free_globals(&mut cs);
        primitives::globals::normalize_global_table(&mut cs);
        let aliases: Vec<_> = cs[0]
            .global_inits
            .iter()
            .filter_map(|g| {
                if let vybe_runtime::chunk::ConstExpr::RefFunc(i) = g.init {
                    Some((g.name.clone(), i))
                } else {
                    None
                }
            })
            .collect();
        for (name, i) in aliases {
            let mut f = Chunk::new(format!("binding-factory:{name}"));
            primitives::functions::emit_ref_func(&mut f, i, 0, 0);
            f.emit_op(Op::RETURN, 0);
            cs.push(f);
        }
        cs
    })
}
fn setup(trace: Arc<Mutex<Vec<(i32, i32)>>>) -> VM {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    vm.register_host_fn(
        "test:sort",
        "compare",
        Box::new(move |_, args| {
            let a = args[args.len() - 2].as_f64() as i32;
            let b = args[args.len() - 1].as_f64() as i32;
            trace.lock().unwrap().push((a, b));
            Value::I32(a - b)
        }),
    );
    vm
}
fn array(v: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(
        vybe_runtime::value::Object::new_array(v),
    ))
}
fn values(v: &Value) -> Vec<Value> {
    let Value::Object(o) = v else { panic!() };
    let o = o.lock().unwrap();
    let vybe_runtime::value::ObjectKind::Array(v) = &o.kind else {
        panic!()
    };
    v.clone()
}
#[test]
fn values_callbacks_and_sizes() {
    let mut checks = 0;
    for v in [
        Value::I32(-17),
        Value::I32(0),
        Value::I32(17),
        Value::F64(-0.0),
        Value::F64(0.5),
        Value::F64(f64::NAN),
        Value::F64(f64::INFINITY),
        Value::F64(f64::NEG_INFINITY),
        Value::Bool(false),
        Value::Bool(true),
        Value::Null,
        Value::Undefined,
        Value::String("1".into()),
        Value::bigint_i64(1),
    ] {
        let mut outcomes = vec![];
        for fast in [false, true] {
            let mut vm = setup(Default::default());
            let f = vm.run(program(fast, 0, 1)).unwrap();
            let out = vm.invoke(&f, &[v.clone()]);
            outcomes.push(out.map_err(|e| e.message.clone()));
        }
        assert_eq!(outcomes[0], outcomes[1]);
        checks += 1;
    }
    for kind in 1..5 {
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
                let mut outcomes = vec![];
                for fast in [false, true] {
                    let trace = Arc::new(Mutex::new(vec![]));
                    let mut vm = setup(trace.clone());
                    let f = vm.run(program(fast, kind, 1)).unwrap();
                    let a = array(input.clone());
                    let out = vm.invoke(&f, &[a.clone()]).unwrap();
                    assert_eq!(out, a);
                    assert_eq!(values(&a), expected);
                    outcomes.push(trace.lock().unwrap().clone());
                    checks += 2;
                }
                assert_eq!(outcomes[0], outcomes[1]);
                checks += 1;
            }
        }
    }
    let mut rows = vec![];
    for kind in 0..5 {
        let mut sizes = vec![];
        for fast in [false, true] {
            let cs = program(fast, kind, 1);
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
            if std::env::var_os("VYBE_ZERO_EXPORT").is_some() {
                native_export::export_chunks(&cs, &format!("sort-zero-{kind}-{fast}"));
            }
            sizes.push(serde_json::json!({"bodies":cs.iter().map(|c|c.code.len()).sum::<usize>(),"locals":cs.iter().map(|c|c.local_count as usize).sum::<usize>(),"native":bytes.len()}));
        }
        assert!(sizes[1]["bodies"].as_u64() < sizes[0]["bodies"].as_u64());
        assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
        rows.push(serde_json::json!({"kind":kind,"sizes":sizes}));
    }
    let result = serde_json::json!({"checks":checks,"rows":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../SORT_ZERO_GATES.json"),
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
    for (kind, len, ordered, control) in [
        (0, 64, false, true),
        (0, 64, false, false),
        (1, 32, false, false),
        (1, 256, false, false),
        (1, 256, true, false),
        (2, 256, false, false),
        (3, 256, false, false),
    ] {
        // No trace collection in timing; callback arithmetic is identical for both.
        let mut vm = VM::new();
        primitives::platforms::register_platforms_all(&mut vm);
        vm.register_host_fn(
            "test:sort",
            "compare",
            Box::new(|_, a| Value::F64(a[a.len() - 2].as_f64() - a[a.len() - 1].as_f64())),
        );
        let old = vm
            .run(program(false, kind, if kind == 0 { 64 } else { 1 }))
            .unwrap();
        let new = if control {
            old.clone()
        } else {
            vm.run(program(true, kind, if kind == 0 { 64 } else { 1 }))
                .unwrap()
        };
        let fs = [old, new];
        let input: Vec<Value> = (0..len)
            .map(|i| Value::I32(if ordered { i } else { (i * 17 + 3) % 101 }))
            .collect();
        let call = |vm: &mut VM, f: &Value| {
            vm.invoke(
                f,
                &[if kind == 0 {
                    Value::I32(17)
                } else {
                    array(input.clone())
                }],
            )
            .unwrap()
        };
        for f in &fs {
            for _ in 0..8 {
                std::hint::black_box(call(&mut vm, f));
            }
        }
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let t = Instant::now();
                let mut n = 0;
                while t.elapsed() < Duration::from_millis(500) || n < 8 {
                    std::hint::black_box(call(&mut vm, &fs[side]));
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
        rows.push(serde_json::json!({"kind":kind,"len":len,"ordered":ordered,"control":control,"median_speedup":ratios[3],"seconds_per_call":samples}));
    }
    let result = serde_json::json!({"same_vm":true,"pairs":7,"minimum_sample_ms":500,"minimum_calls":8,"warmups":8,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../SORT_ZERO_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
