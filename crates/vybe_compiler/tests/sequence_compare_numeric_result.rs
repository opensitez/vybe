use vybe_compiler::primitives::{self, collections};
use vybe_runtime::value::Object;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/sequence_compare_dynamic_result.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("sequence-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("sequence");
    c.arity = 2;
    c.param_count = 2;
    c.local_count = 2;
    for s in 0..2 {
        c.emit_op_u16(Op::LOCAL_GET, s, 0);
    }
    let mut cs = vec![root, c];
    if fast {
        collections::emit_sequence_compare(&mut cs, 1, 0);
    } else {
        before::emit_sequence_compare(&mut cs, 1, 0);
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
#[test]
fn sequence_result_preserves_numeric_prefix_and_type_behavior() {
    let mut checks = 0;
    let mut workers = [worker(false), worker(true)];
    for l in 0..17usize {
        for r in 0..17usize {
            for difference in [None, Some(0usize), Some(7), Some(15)] {
                let left: Vec<_> = (0..l).map(|i| Value::I32(i as i32)).collect();
                let right: Vec<_> = (0..r)
                    .map(|i| Value::I32(i as i32 + if difference == Some(i) { 1 } else { 0 }))
                    .collect();
                let expected = if difference.is_some_and(|i| i < l.min(r)) {
                    -1.0
                } else {
                    l.cmp(&r) as i8 as f64
                };
                let args = [array(left), array(right)];
                let out = workers
                    .each_mut()
                    .map(|(vm, f)| vm.invoke(f, &args).unwrap().as_f64());
                assert_eq!(out[0], out[1]);
                assert_eq!(out[1], expected);
                checks += 1;
            }
        }
    }
    let values = [
        Value::Null,
        Value::Undefined,
        Value::Bool(false),
        Value::Bool(true),
        Value::I32(0),
        Value::F64(-0.0),
        Value::F64(f64::NAN),
        Value::F64(f64::INFINITY),
        Value::String("".into()),
        Value::String("a".into()),
        Value::String("é😀".into()),
        Value::bigint_i64(1),
    ];
    for a in &values {
        for b in &values {
            let args = [array(vec![a.clone()]), array(vec![b.clone()])];
            let mut results = vec![];
            for fast in [false, true] {
                let (mut vm, f) = worker(fast);
                results.push(
                    vm.invoke(&f, &args)
                        .map(|v| v.as_f64())
                        .map_err(|e| e.message),
                );
            }
            assert_eq!(results[0], results[1]);
            checks += 1;
        }
    }
    eprintln!("Verified {checks} numeric/prefix/mixed-type sequence comparisons");
}
#[test]
fn native_size_and_exports() {
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
        if std::env::var_os("VYBE_PRIMITIVE_SEQUENCE_EXPORT").is_some() {
            native_export::export_chunks(chunks, &format!("sequence-{side}"));
        }
        sizes.push(serde_json::json!({"body":chunks[1].code.len(),"internal":chunks.iter().map(|c|c.code.len()).sum::<usize>(),"native":bytes.len(),"locals":chunks[1].local_count,"chunks":chunks.len()}));
    }
    assert!(cs[1][1].code.len() < cs[0][1].code.len());
    assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
    eprintln!("{sizes:?}");
    if std::env::var_os("VYBE_PRIMITIVE_SEQUENCE_EXPORT").is_some() {
        std::fs::write(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../SEQUENCE_GATES.json"),
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
    for (length, difference) in [
        (0usize, None),
        (8, None),
        (4096, Some(0usize)),
        (4096, None),
    ] {
        let mut workers = [worker(false), worker(true)];
        let args = [
            array((0..length).map(|i| Value::I32(i as i32)).collect()),
            array(
                (0..length)
                    .map(|i| Value::I32(i as i32 + if difference == Some(i) { 1 } else { 0 }))
                    .collect(),
            ),
        ];
        let expected = if difference.is_some() { -1.0 } else { 0.0 };
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let (vm, f) = &mut workers[side];
                let start = Instant::now();
                let mut n = 0;
                while start.elapsed() < Duration::from_millis(100) || n < 8 {
                    assert_eq!(
                        std::hint::black_box(vm.invoke(f, &args).unwrap()).as_f64(),
                        expected
                    );
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
        rows.push(serde_json::json!({"length":length,"difference":difference,"median_speedup":ratios[3],"seconds_per_call":samples}));
    }
    let result = serde_json::json!({"profile":"opt0 assertions and overflow checks","pairs":7,"minimum_sample_ms":100,"minimum_calls_per_sample":8,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../SEQUENCE_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
