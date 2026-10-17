use vybe_compiler::primitives;
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/reverse_range_before.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("reverse-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut imports = Chunk::new("imports");
    let mut c = if fast {
        primitives::polyfills::build_runtime_helpers_for_exports(
            &mut imports,
            &["__stdlib_array_reverse_range"],
        )
        .chunks
        .remove(0)
    } else {
        before::build_array_reverse_range(&mut imports)
    };
    c.name = "reverse".into();
    c.param_count = c.arity;
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
fn contents(v: &Value) -> Vec<Value> {
    let Value::Object(o) = v else { panic!() };
    let o = o.lock().unwrap();
    let vybe_runtime::value::ObjectKind::Array(v) = &o.kind else {
        panic!()
    };
    v.clone()
}
#[test]
fn bounds_mutations_and_sizes() {
    let mut checks = 0;
    let mut numeric_workers = [worker(false), worker(true)];
    for len in 0..33 {
        for start in 0..=len {
            for count in 0..=len - start {
                let expected: Vec<_> = (0..len).map(Value::I32).collect();
                let mut expected = expected;
                expected[start as usize..(start + count) as usize].reverse();
                for (vm, f) in &mut numeric_workers {
                    let input = array((0..len).map(Value::I32).collect());
                    assert_eq!(
                        vm.invoke(&f, &[input.clone(), Value::I32(start), Value::I32(count)])
                            .unwrap(),
                        Value::Null
                    );
                    assert_eq!(contents(&input), expected);
                    checks += 1;
                }
            }
        }
    }
    for (start, count) in [
        (Value::Bool(false), Value::I32(4)),
        (Value::Null, Value::I32(4)),
        (Value::I32(0), Value::String("2".into())),
        (Value::String("9".into()), Value::I32(0)),
        (Value::I32(-1), Value::I32(3)),
        (Value::F64(0.5), Value::F64(3.5)),
        (Value::F64(f64::NAN), Value::I32(3)),
        (Value::F64(f64::INFINITY), Value::I32(0)),
        (Value::I32(0), Value::bigint_i64(3)),
        (Value::Undefined, Value::I32(3)),
    ] {
        let mut outcomes = vec![];
        for fast in [false, true] {
            let (mut vm, f) = worker(fast);
            let a = array((0..8).map(Value::I32).collect());
            let result = vm.invoke(&f, &[a.clone(), start.clone(), count.clone()]);
            outcomes.push((result.is_ok(), contents(&a)));
        }
        assert_eq!(outcomes[0], outcomes[1]);
        checks += 1;
    }
    let mut rows = vec![];
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
        if std::env::var_os("VYBE_REVERSE_RANGE_EXPORT").is_some() {
            native_export::export_chunks(&cs, &format!("reverse-range-{fast}"));
        }
        rows.push(serde_json::json!({"body":cs[1].code.len(),"locals":cs[1].local_count,"all_bodies":cs.iter().map(|c|c.code.len()).sum::<usize>(),"native":bytes.len(),"chunks":cs.iter().map(|c|serde_json::json!({"name":c.name,"body":c.code.len()})).collect::<Vec<_>>() }));
    }
    assert!(rows[1]["all_bodies"].as_u64() < rows[0]["all_bodies"].as_u64());
    assert!(rows[1]["native"].as_u64() < rows[0]["native"].as_u64());
    let result = serde_json::json!({"checks":checks,"sizes":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../REVERSE_RANGE_GATES.json"),
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
    for len in [0, 8, 256, 4096] {
        let mut ws = [worker(false), worker(true)];
        let input = array((0..len).map(Value::I32).collect());
        let args = [input, Value::I32(0), Value::I32(len)];
        for w in &mut ws {
            for _ in 0..8 {
                std::hint::black_box(w.0.invoke(&w.1, &args).unwrap());
            }
        }
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let t = Instant::now();
                let mut n = 0;
                let f = ws[side].1.clone();
                while t.elapsed() < Duration::from_millis(200) || n < 8 {
                    std::hint::black_box(ws[side].0.invoke(&f, &args).unwrap());
                    n += 1;
                }
                samples[side].push(t.elapsed().as_secs_f64() / n as f64);
            }
        }
        let mut r: Vec<_> = samples[0]
            .iter()
            .zip(&samples[1])
            .map(|(a, b)| a / b)
            .collect();
        r.sort_by(f64::total_cmp);
        rows.push(serde_json::json!({"len":len,"median_speedup":r[3],"seconds_per_call":samples}));
    }
    let result = serde_json::json!({"pairs":7,"minimum_sample_ms":200,"minimum_calls":8,"warmups":8,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../REVERSE_RANGE_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}

#[test]
fn private_math_links_once_with_wide_function_indices() {
    let mut imports = Chunk::new("imports");
    let mut cs = primitives::polyfills::build_runtime_helpers_for_exports(
        &mut imports, &["__stdlib_array_reverse_range"],
    ).chunks;
    cs.extend((1..=u16::MAX as usize).map(|i| Chunk::new(format!("padding-{i}"))));
    primitives::bundle::finalize_with_runtime_helpers(&mut cs);
    for name in ["primitive.array_range_add", "primitive.array_range_lt"] {
        let targets: Vec<_> = cs.iter().enumerate().filter(|(_, c)| c.name == name).map(|(i, _)| i).collect();
        assert_eq!(targets.len(), 1);
        assert!(targets[0] > u16::MAX as usize);
        let edges: Vec<_> = cs[0].direct_defined_calls.values().filter(|e| e.target as usize == targets[0]).collect();
        assert_eq!(edges.len(), if name.ends_with("add") { 4 } else { 1 });
        assert!(edges.iter().all(|e| !e.source && e.params == "externref,externref" && e.results == "externref"));
    }
    let len = cs.len();
    let edges = cs[0].direct_defined_calls.clone();
    primitives::bundle::finalize_with_runtime_helpers(&mut cs);
    assert_eq!(cs.len(), len);
    assert_eq!(format!("{:?}", cs[0].direct_defined_calls), format!("{edges:?}"));
}
