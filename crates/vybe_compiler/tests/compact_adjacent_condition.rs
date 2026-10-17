use vybe_compiler::primitives::{self, collections};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/compact_adjacent_truthiness.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("compact-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("compact");
    c.arity = 1;
    c.param_count = 1;
    c.local_count = 1;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let mut cs = vec![root, c];
    if fast {
        collections::emit_compact_adjacent(&mut cs, 1, 0);
    } else {
        before::emit_compact_adjacent(&mut cs, 1, 0);
    }
    cs[1].emit_op(Op::RETURN, 0);
    primitives::bundle::finalize_with_runtime_helpers(&mut cs);
    primitives::globals::declare_free_globals(&mut cs);
    primitives::globals::normalize_global_table(&mut cs);
    let bindings: Vec<_> = cs[0]
        .global_inits
        .iter()
        .filter_map(|init| {
            if let vybe_runtime::chunk::ConstExpr::RefFunc(index) = init.init {
                Some((init.name.clone(), index))
            } else {
                None
            }
        })
        .collect();
    for (name, index) in bindings {
        let mut f = Chunk::new(format!("binding-factory:{name}"));
        f.emit_op_u16(Op::REF_FUNC, index as u16, 0);
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
fn array(values: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)))
}
fn items(v: &Value) -> Vec<Value> {
    let Value::Object(o) = v else { panic!("array") };
    let o = o.lock().unwrap();
    let ObjectKind::Array(v) = &o.kind else {
        panic!("array")
    };
    v.clone()
}
fn equal(a: &[Value], b: &[Value]) {
    assert_eq!(a.len(), b.len());
    for (a, b) in a.iter().zip(b) {
        match (a, b) {
            (Value::F64(a), Value::F64(b)) => assert_eq!(a.to_bits(), b.to_bits()),
            (Value::Object(a), Value::Object(b)) => assert!(std::sync::Arc::ptr_eq(a, b)),
            _ => assert_eq!(a, b),
        }
    }
}
#[test]
fn compact_preserves_runs_values_identity_and_input() {
    let object = array(vec![Value::I32(1)]);
    let mut checks = 0;
    for values in [
        vec![],
        vec![Value::I32(1)],
        vec![Value::I32(1); 128],
        (0..128).map(Value::I32).collect(),
        vec![
            Value::I32(1),
            Value::I32(1),
            Value::I32(2),
            Value::I32(2),
            Value::I32(1),
        ],
        vec![
            Value::Null,
            Value::Null,
            Value::Undefined,
            Value::Bool(false),
            Value::Bool(false),
        ],
        vec![
            Value::String("é😀".into()),
            Value::String("é😀".into()),
            Value::String("x".into()),
        ],
        vec![
            Value::F64(-0.0),
            Value::F64(0.0),
            Value::F64(f64::NAN),
            Value::F64(f64::NAN),
        ],
        vec![
            Value::bigint_i64(1),
            Value::bigint_i64(1),
            Value::bigint_i64(2),
        ],
        vec![
            object.clone(),
            object.clone(),
            array(vec![Value::I32(1)]),
            object,
        ],
    ] {
        let input = array(values.clone());
        let mut results = vec![];
        for fast in [false, true] {
            let (mut vm, f) = worker(fast);
            let out = vm.invoke(&f, &[input.clone()]).unwrap();
            let again = vm.invoke(&f, &[input.clone()]).unwrap();
            let (Value::Object(a), Value::Object(b)) = (&out, again) else {
                panic!("array")
            };
            assert!(!std::sync::Arc::ptr_eq(a, &b));
            equal(&items(&input), &values);
            results.push(items(&out));
        }
        equal(&results[0], &results[1]);
        checks += 1;
    }
    for mask in 0..128u32 {
        let values = (0..8)
            .map(|i| Value::I32(((mask >> i) & 1) as i32))
            .collect::<Vec<_>>();
        let mut expected = values.clone();
        expected.dedup();
        let mut workers = [worker(false), worker(true)];
        let input = array(values);
        for (vm, f) in &mut workers {
            let result = vm.invoke(f, &[input.clone()]).unwrap();
            equal(&items(&result), &expected);
        }
        checks += 1;
    }
    eprintln!("Verified {checks} compact run/value/identity cases");
}
#[test]
fn size_and_native_exports() {
    let cs = [program(false), program(true)];
    let mut sizes = vec![];
    for (side, chunks) in cs.iter().enumerate() {
        let bytes = vybe_platform_wasm::write_wasm_with_options(
            chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_name_section: false,
                include_vybe_metadata: false,
            },
        );
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes)
            .unwrap();
        if std::env::var_os("VYBE_PRIMITIVE_COMPACT_EXPORT").is_some() {
            native_export::export_chunks(chunks, &format!("compact-{side}"));
        }
        sizes.push(serde_json::json!({"body":chunks[1].code.len(),"native":bytes.len(),"locals":chunks[1].local_count}));
    }
    assert!(cs[1][1].code.len() < cs[0][1].code.len());
    assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
    eprintln!("{sizes:?}");
    if std::env::var_os("VYBE_PRIMITIVE_COMPACT_EXPORT").is_some() {
        std::fs::write(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../COMPACT_GATES.json"),
            serde_json::to_vec_pretty(&sizes).unwrap(),
        )
        .unwrap();
    }
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::{Duration, Instant};
    let mut rows = vec![];
    for (length, run) in [
        (0usize, 1usize),
        (8, 2),
        (4096, 1),
        (4096, 16),
        (4096, 4096),
    ] {
        let mut workers = [worker(false), worker(true)];
        let input = array((0..length).map(|i| Value::I32((i / run) as i32)).collect());
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let (vm, f) = &mut workers[side];
                let start = Instant::now();
                let mut n = 0;
                while start.elapsed() < Duration::from_millis(100) || n < 8 {
                    std::hint::black_box(vm.invoke(f, &[input.clone()]).unwrap());
                    n += 1;
                }
                samples[side].push(start.elapsed().as_secs_f64() / n as f64);
            }
        }
        let mut ratios: Vec<_> = samples[0]
            .iter()
            .zip(&samples[1])
            .map(|(a, b)| a / b)
            .collect();
        ratios.sort_by(f64::total_cmp);
        rows.push(serde_json::json!({"length":length,"run_length":run,"median_speedup":ratios[3],"seconds_per_call":samples}));
    }
    let result = serde_json::json!({"profile":"opt0 assertions and overflow checks","pairs":7,"minimum_sample_ms":100,"minimum_calls_per_sample":8,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../COMPACT_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
