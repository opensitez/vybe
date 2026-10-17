use vybe_compiler::primitives::{self, ops};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/array_range_generic_add.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool, kind: u8, repeats: usize) -> Vec<Chunk> {
    let mut root = Chunk::new("range-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut imports = Chunk::new("imports");
    let mut c = match (fast, kind) {
        (false, 0) => before::build_array_insert_range(&mut imports),
        (false, 1) => before::build_array_set_range(&mut imports),
        (false, 2) => before::build_array_reverse_range(&mut imports),
        (true, 0..=2) => primitives::polyfills::build_runtime_helpers_for_exports(
            &mut imports,
            &[["__stdlib_array_insert_range", "__stdlib_array_set_range", "__stdlib_array_reverse_range"][kind as usize]],
        ).chunks.remove(0),
        (_, 3) => {
            let mut c = Chunk::new("addition");
            c.arity = 2;
            c.local_count = 2;
            for i in 0..repeats {
                c.emit_op_u16(Op::LOCAL_GET, 0, 0);
                c.emit_op_u16(Op::LOCAL_GET, 1, 0);
                if fast {
                    ops::emit_dyn_add_number_rhs_into(&mut imports, &mut c, 0)
                } else {
                    ops::emit_dyn_add_into(&mut imports, &mut c, 0)
                }
                if i + 1 < repeats {
                    c.emit_op(Op::DROP, 0);
                }
            }
            c.emit_op(Op::RETURN, 0);
            c
        }
        _ => unreachable!(),
    };
    c.name = "range-worker".into();
    c.param_count = c.arity;
    let mut cs = vec![root, c];
    primitives::bundle::finalize_with_runtime_helpers(&mut cs);
    primitives::globals::declare_free_globals(&mut cs);
    primitives::globals::normalize_global_table(&mut cs);
    let bindings: Vec<_> = cs[0]
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
    for (name, i) in bindings {
        let mut f = Chunk::new(format!("binding-factory:{name}"));
        f.emit_op_u16(Op::REF_FUNC, i as u16, 0);
        f.emit(0, 0);
        f.emit_op(Op::RETURN, 0);
        cs.push(f);
    }
    cs
}
fn worker(fast: bool, kind: u8, repeats: usize) -> (VM, Value) {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(program(fast, kind, repeats)).unwrap();
    (vm, f)
}
fn array(values: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(
        vybe_runtime::value::Object::new_array(values),
    ))
}
fn contents(v: &Value) -> Vec<Value> {
    if let Value::Object(o) = v {
        if let vybe_runtime::value::ObjectKind::Array(a) = &o.lock().unwrap().kind {
            return a.clone();
        }
    }
    panic!("array expected")
}
#[test]
fn numeric_rhs_preserves_legacy_addition_and_range_effects() {
    let mut checks = 0;
    let left = [
        Value::I32(3),
        Value::F64(-0.0),
        Value::F64(f64::NAN),
        Value::F64(f64::INFINITY),
        Value::String("1".into()),
        Value::String("é😀".into()),
        Value::String("".into()),
        Value::Null,
        Value::Undefined,
        Value::Bool(true),
        Value::Bool(false),
        Value::bigint_i64(3),
        array(vec![Value::I32(3)]),
    ];
    let right = [
        Value::I32(0),
        Value::I32(1),
        Value::I32(i32::MAX),
        Value::F64(-1.0),
        Value::F64(-0.0),
        Value::F64(0.25),
        Value::F64(f64::NAN),
        Value::F64(f64::INFINITY),
    ];
    let mut add = [worker(false, 3, 1), worker(true, 3, 1)];
    for a in &left {
        for b in &right {
            let out = add
                .each_mut()
                .map(|(vm, f)| vm.invoke(f, &[a.clone(), b.clone()]).map_err(|e| e.message));
            assert_eq!(out[0], out[1], "a={a:?} b={b:?}");
            checks += 1;
        }
    }
    for kind in 0..3 {
        let mut workers = [worker(false, kind, 1), worker(true, kind, 1)];
        for len in [0, 1, 8, 32] {
            for start in [0, 1, 3] {
                for n in [0, 1, 4] {
                    let original: Vec<_> = (0..len).map(Value::I32).collect();
                    let source: Vec<_> = (0..n).map(|i| Value::I32(100 + i)).collect();
                    let mut outcomes = vec![];
                    for (vm, f) in &mut workers {
                        let arr = array(original.clone());
                        let rhs = if kind == 2 {
                            Value::I32(n)
                        } else {
                            array(source.clone())
                        };
                        let result = vm
                            .invoke(f, &[arr.clone(), Value::I32(start), rhs])
                            .map_err(|e| e.message);
                        outcomes.push((result, contents(&arr)));
                    }
                    assert_eq!(
                        outcomes[0], outcomes[1],
                        "kind={kind} len={len} start={start} n={n}"
                    );
                    if start + n <= len {
                        let mut expected = original.clone();
                        match kind {
                            0 => {
                                expected.splice(start as usize..start as usize, source.clone());
                            }
                            1 => expected[start as usize..(start + n) as usize]
                                .clone_from_slice(&source),
                            2 => expected[start as usize..(start + n) as usize].reverse(),
                            _ => unreachable!(),
                        };
                        assert_eq!(outcomes[1], (Ok(Value::Null), expected));
                    }
                    checks += 1;
                }
            }
        }
        if kind < 2 {
            for index in [
                Value::String("1".into()),
                Value::Null,
                Value::Bool(true),
                Value::F64(0.25),
                Value::F64(f64::NAN),
            ] {
                let mut outcomes = vec![];
                for (vm, f) in &mut workers {
                    let arr = array(vec![Value::I32(0); 8]);
                    let result = vm
                        .invoke(
                            f,
                            &[arr.clone(), index.clone(), array(vec![Value::I32(9); 3])],
                        )
                        .map_err(|e| e.message);
                    outcomes.push((result, contents(&arr)));
                }
                assert_eq!(outcomes[0], outcomes[1]);
                checks += 1;
            }
        }
        // Aliasing is observable: retain the sequential reads/writes.
        for start in [0, 1] {
            let mut outcomes = vec![];
            for (vm, f) in &mut workers {
                let arr = array(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
                let rhs = if kind == 2 {
                    Value::I32(2)
                } else {
                    arr.clone()
                };
                let result = vm
                    .invoke(f, &[arr.clone(), Value::I32(start), rhs])
                    .map_err(|e| e.message);
                outcomes.push((result, contents(&arr)));
            }
            assert_eq!(outcomes[0], outcomes[1]);
            checks += 1;
        }
    }
    eprintln!("Verified {checks} legacy addition/range outcome and mutation cases");
}
#[test]
fn size_and_native_exports() {
    let mut rows = vec![];
    for kind in 0..4 {
        let mut sizes = vec![];
        for fast in [false, true] {
            let cs = program(fast, kind, 1);
            let c = &cs[1];
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
            if std::env::var_os("VYBE_PRIMITIVE_RANGE_ADD_EXPORT").is_some() {
                native_export::export_chunks(&cs, &format!("range-add-{kind}-{}", u8::from(fast)));
            }
            sizes.push(serde_json::json!({"body":c.code.len(),"locals":c.local_count,"native":bytes.len()}));
        }
        if kind == 0 || kind == 2 {
            assert_eq!(sizes[0], sizes[1]);
        } else {
            assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
        }
        rows.push(serde_json::json!({"kind":kind,"sizes":sizes}));
    }
    if std::env::var_os("VYBE_PRIMITIVE_RANGE_ADD_EXPORT").is_some() {
        std::fs::write(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../RANGE_ADD_GATES.json"),
            serde_json::to_vec_pretty(&rows).unwrap(),
        )
        .unwrap();
    }
    eprintln!("{rows:?}");
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::{Duration, Instant};
    let mut rows = vec![];
    for (label, kind, len, n) in [
        ("addition-number", 3, 0, 0),
        ("insert-8", 0, 8, 8),
        ("set-4096", 1, 4096, 4096),
        ("reverse-4096", 2, 4096, 4096),
    ] {
        let mut workers = [worker(false, kind, 64), worker(true, kind, 64)];
        let mut samples = [vec![], vec![]];
        let original: Vec<_> = (0..len).map(Value::I32).collect();
        let source = array(vec![Value::I32(7); n as usize]);
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let (vm, f) = &mut workers[side];
                let fixed = array(original.clone());
                let start = Instant::now();
                let mut count = 0;
                while start.elapsed() < Duration::from_millis(100) || count < 8 {
                    let out = if kind == 3 {
                        vm.invoke(f, &[Value::F64(3.0), Value::I32(1)])
                    } else {
                        let arr = if kind == 0 {
                            array(original.clone())
                        } else {
                            fixed.clone()
                        };
                        vm.invoke(
                            f,
                            &[
                                arr,
                                Value::I32(0),
                                if kind == 2 {
                                    Value::I32(n)
                                } else {
                                    source.clone()
                                },
                            ],
                        )
                    };
                    std::hint::black_box(out.unwrap());
                    count += 1;
                }
                samples[side].push(start.elapsed().as_secs_f64() / count as f64);
            }
        }
        let mut ratios: Vec<_> = samples[0]
            .iter()
            .zip(&samples[1])
            .map(|(a, b)| a / b)
            .collect();
        ratios.sort_by(f64::total_cmp);
        rows.push(
            serde_json::json!({"case":label,"median_speedup":ratios[3],"seconds_per_call":samples}),
        );
    }
    let result = serde_json::json!({"pairs":7,"minimum_sample_ms":100,"minimum_calls":8,"addition_operations_per_call":64,"insert_includes_fresh_destination":true,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../RANGE_ADD_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
