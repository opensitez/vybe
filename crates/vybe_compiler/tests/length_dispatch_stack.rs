use vybe_compiler::primitives::{self, collections};
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/length_dispatch_before.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool, kind: usize, repeats: usize) -> Vec<Chunk> {
    let mut root = Chunk::new("length-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = if kind == 0 || kind == 5 {
        let mut c = Chunk::new("length");
        c.arity = 1;
        c.param_count = 1;
        c.local_count = 1;
        for n in 0..repeats {
            if kind == 5 {
                c.emit_i32_const(17, 0);
            }
            c.emit_op_u16(Op::LOCAL_GET, 0, 0);
            if fast {
                collections::emit_len_into(&mut root, &mut c, 0);
            } else {
                before::emit_len_into(&mut root, &mut c, 0);
            }
            if kind == 5 {
                c.emit_op(Op::I32_ADD, 0);
            }
            if n + 1 < repeats {
                c.emit_op(Op::DROP, 0);
            }
        }
        c.emit_op(Op::RETURN, 0);
        c
    } else {
        match (kind, fast) {
            (1, true) => collections::build_sum(&mut root),
            (1, false) => before::build_sum(&mut root),
            (2, true) => collections::build_min(&mut root),
            (2, false) => before::build_min(&mut root),
            (3, true) => collections::build_max(&mut root),
            (3, false) => before::build_max(&mut root),
            (_, true) => collections::build_isempty(&mut root),
            (_, false) => before::build_isempty(&mut root),
        }
    };
    c.name = "length".into();
    let mut cs = vec![root, c];
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
        f.emit_op_u16(Op::REF_FUNC, i as u16, 0);
        f.emit(0, 0);
        f.emit_op(Op::RETURN, 0);
        cs.push(f);
    }
    cs
}
fn worker(fast: bool, kind: usize, repeats: usize) -> (VM, Value) {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(program(fast, kind, repeats)).unwrap();
    (vm, f)
}
fn array(v: Vec<Value>) -> Value {
    object(vybe_runtime::value::Object::new_array(v))
}
fn object(o: vybe_runtime::value::Object) -> Value {
    Value::Object(vybe_runtime::heap::alloc(o))
}
fn run(w: &mut (VM, Value), v: &Value) -> Value {
    w.0.invoke(&w.1, &[v.clone()]).unwrap()
}
#[test]
fn values_and_sizes() {
    let mut checks = 0;
    let mut rows = vec![];
    for kind in 0..6 {
        let mut ws = [worker(false, kind, 1), worker(true, kind, 1)];
        if kind == 0 {
            let mut cases = vec![
                (Value::Null, Value::Null),
                (Value::Undefined, Value::Null),
                (Value::I32(17), Value::Null),
                (Value::Bool(false), Value::Null),
                (object(vybe_runtime::value::Object::new()), Value::Null),
            ];
            for s in ["", "abc", "é漢😀", "\0"] {
                cases.push((
                    Value::String(s.into()),
                    Value::I32(s.encode_utf16().count() as i32),
                ));
            }
            for n in [0, 1, 17, 64] {
                cases.push((array((0..n).map(Value::I32).collect()), Value::I32(n)));
                let mut o = vybe_runtime::value::Object::new();
                o.properties.insert("length".into(), Value::I32(n));
                cases.push((object(o), Value::I32(n)));
            }
            for (input, expected) in cases {
                for w in &mut ws {
                    assert_eq!(run(w, &input), expected);
                    checks += 1;
                }
            }
        } else if kind == 5 {
            for input in [
                array(vec![]),
                array(vec![Value::I32(9); 17]),
                Value::String("é漢😀".into()),
            ] {
                let a = run(&mut ws[0], &input);
                let b = run(&mut ws[1], &input);
                assert_eq!(a, b);
                let len = match &input {
                    Value::String(s) => s.encode_utf16().count(),
                    Value::Object(o) => {
                        let o = o.lock().unwrap();
                        if let vybe_runtime::value::ObjectKind::Array(v) = &o.kind {
                            v.len()
                        } else {
                            unreachable!()
                        }
                    }
                    _ => unreachable!(),
                };
                assert_eq!(b, Value::I32(17 + len as i32));
                checks += 1;
            }
        } else {
            for n in 0..96 {
                let vals: Vec<i32> = (0..n).map(|i| (i * 17 % 53) - 26).collect();
                let input = array(vals.iter().copied().map(Value::I32).collect());
                let expected = match kind {
                    1 => Some(vals.iter().sum::<i32>() as f64),
                    2 => vals.iter().min().map(|v| *v as f64),
                    3 => vals.iter().max().map(|v| *v as f64),
                    _ => Some(f64::from(n == 0)),
                };
                let a = run(&mut ws[0], &input);
                let b = run(&mut ws[1], &input);
                assert_eq!(a, b);
                if let Some(expected) = expected {
                    assert_eq!(b.as_f64(), expected);
                } else {
                    assert_eq!(b, Value::Undefined);
                }
                checks += 1;
            }
        }
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
            if std::env::var_os("VYBE_LENGTH_EXPORT").is_some() {
                native_export::export_chunks(&cs, &format!("length-{kind}-{fast}"));
            }
            sizes.push(serde_json::json!({"body":cs[1].code.len(),"locals":cs[1].local_count,"native":bytes.len()}));
        }
        assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
        assert!(sizes[1]["locals"].as_u64() < sizes[0]["locals"].as_u64());
        rows.push(serde_json::json!({"kind":kind,"sizes":sizes}));
    }
    let result = serde_json::json!({"checks":checks,"rows":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../LENGTH_STACK_GATES.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
fn paired_worker(kind: usize, repeats: usize, identical: bool) -> (VM, [Value; 2]) {
    let old = program(false, kind, repeats).remove(1);
    let new = program(!identical, kind, repeats).remove(1);
    let mut root = Chunk::new("paired-length-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    primitives::functions::emit_ref_func(&mut root, 2, 0, 0);
    let mut cs = vec![root, old, new];
    collections::emit_array_pair(&mut cs, 0, 0);
    cs[0].emit_op(Op::RETURN, 0);
    primitives::bundle::finalize_with_runtime_helpers(&mut cs);
    primitives::globals::declare_free_globals(&mut cs);
    primitives::globals::normalize_global_table(&mut cs);
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let api = vm.run(cs).unwrap();
    let Value::Object(o) = api else {
        panic!("pair")
    };
    let o = o.lock().unwrap();
    let vybe_runtime::value::ObjectKind::Array(values) = &o.kind else {
        panic!("pair")
    };
    (vm, [values[0].clone(), values[1].clone()])
}

#[test]
#[ignore]
fn debug_pairs() {
    use std::time::{Duration, Instant};
    let mut rows = vec![];
    for (kind, label, input) in [
        (0, "length-array-64-calls", array(vec![Value::I32(0); 64])),
        (
            0,
            "length-string-64-calls",
            Value::String("é漢😀".repeat(16).into()),
        ),
        (1, "sum-empty", array(vec![])),
        (1, "sum-64", array((0..64).map(Value::I32).collect())),
        (2, "min-64", array((0..64).rev().map(Value::I32).collect())),
        (3, "max-64", array((0..64).map(Value::I32).collect())),
        (
            3,
            "max-64-identical-control",
            array((0..64).map(Value::I32).collect()),
        ),
        (4, "is-empty", array(vec![])),
    ] {
        let (mut vm, functions) = paired_worker(
            kind,
            if kind == 0 { 64 } else { 1 },
            label == "max-64-identical-control",
        );
        for f in &functions {
            for _ in 0..8 {
                std::hint::black_box(vm.invoke(f, &[input.clone()]).unwrap());
            }
        }
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let t = Instant::now();
                let mut n = 0;
                while t.elapsed() < Duration::from_millis(200) || n < 8 {
                    std::hint::black_box(vm.invoke(&functions[side], &[input.clone()]).unwrap());
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
        rows.push(
            serde_json::json!({"case":label,"median_speedup":ratios[3],"seconds_per_call":samples}),
        );
    }
    let result = serde_json::json!({"worker_scope":"old/new in one VM sharing platform/import state","pairs":7,"minimum_sample_ms":200,"minimum_calls":8,"warmups":8,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../LENGTH_STACK_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
