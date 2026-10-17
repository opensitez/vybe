use vybe_compiler::primitives::{
    self,
    strings::{self, SplitOptions},
};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/split_limit_generic_clamp.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool, pieces: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("split-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("split");
    c.arity = 3;
    c.param_count = 3;
    c.local_count = 3;
    for i in 0..3 {
        c.emit_op_u16(Op::LOCAL_GET, i, 0);
    }
    let mut cs = vec![root, c];
    let opts = if pieces {
        SplitOptions::max_pieces()
    } else {
        SplitOptions::max_splits()
    };
    if fast {
        strings::emit_split_limit(&mut cs, 1, 3, opts, 0);
    } else {
        before::emit_split_limit(&mut cs, 1, 3, opts, 0);
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
fn worker(fast: bool, pieces: bool) -> (VM, Value) {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(program(fast, pieces)).unwrap();
    (vm, f)
}
fn items(v: Value) -> Vec<Value> {
    let Value::Object(o) = v else { panic!("array") };
    let o = o.lock().unwrap();
    let ObjectKind::Array(v) = &o.kind else {
        panic!("array")
    };
    v.clone()
}
#[test]
fn limited_split_preserves_empty_unicode_limits_and_errors() {
    let mut checks = 0;
    for pieces in [false, true] {
        let mut workers = [worker(false, pieces), worker(true, pieces)];
        for input in ["", "a", "a,b,c", ",a,,b,", "a::b::c", "é😀é😀尾", "a\0b\0c"] {
            for sep in [",", "::", "", "é", "😀", "\0"] {
                for limit in [
                    -100.0,
                    -3.0,
                    -1.5,
                    -1.0,
                    -0.0,
                    0.0,
                    0.5,
                    1.0,
                    1.5,
                    2.0,
                    3.0,
                    100.0,
                    f64::NAN,
                    f64::INFINITY,
                    f64::NEG_INFINITY,
                ] {
                    let args = [
                        Value::String(input.into()),
                        Value::String(sep.into()),
                        Value::F64(limit),
                    ];
                    let out = workers
                        .each_mut()
                        .map(|(vm, f)| vm.invoke(f, &args).map(items).map_err(|e| e.message));
                    assert_eq!(
                        out[0], out[1],
                        "pieces={pieces} input={input:?} sep={sep:?} limit={limit}"
                    );
                    checks += 1;
                }
            }
        }
        for value in [
            Value::Null,
            Value::Undefined,
            Value::Bool(true),
            Value::String("2".into()),
            Value::Object(vybe_runtime::heap::alloc(Object::new())),
            Value::bigint_i64(2),
        ] {
            let mut out = vec![];
            for fast in [false, true] {
                let (mut vm, f) = worker(fast, pieces);
                out.push(
                    vm.invoke(
                        &f,
                        &[
                            Value::String("a,b,c".into()),
                            Value::String(",".into()),
                            value.clone(),
                        ],
                    )
                    .map(items)
                    .map_err(|e| e.message),
                );
            }
            assert_eq!(out[0], out[1]);
            checks += 1;
        }
    }
    eprintln!("Verified {checks} split cases");
}
#[test]
fn code_size_and_native_exports() {
    let mut rows = vec![];
    for pieces in [false, true] {
        let cs = [program(false, pieces), program(true, pieces)];
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
            if std::env::var_os("VYBE_PRIMITIVE_SPLIT_EXPORT").is_some() {
                native_export::export_chunks(chunks, &format!("split-{pieces}-{side}"));
            }
            sizes.push(serde_json::json!({"internal":chunks[1].code.len(),"native":bytes.len(),"locals":chunks[1].local_count}));
        }
        assert!(cs[1][1].code.len() < cs[0][1].code.len());
        assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
        rows.push(serde_json::json!({"pieces":pieces,"sizes":sizes}));
    }
    eprintln!("{rows:?}");
    if std::env::var_os("VYBE_PRIMITIVE_SPLIT_EXPORT").is_some() {
        std::fs::write(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../SPLIT_GATES.json"),
            serde_json::to_vec_pretty(&rows).unwrap(),
        )
        .unwrap();
    }
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::{Duration, Instant};
    let mut rows = vec![];
    for pieces in [false, true] {
        for (input, limit) in [
            ("a,b,c".to_string(), 2.0),
            ("a,b,c".to_string(), -1.0),
            ("a,".repeat(4096), 2.0),
        ] {
            let mut workers = [worker(false, pieces), worker(true, pieces)];
            let args = [
                Value::String(input.clone().into()),
                Value::String(",".into()),
                Value::F64(limit),
            ];
            let mut samples = [vec![], vec![]];
            for pair in 0..7 {
                for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                    let (vm, f) = &mut workers[side];
                    let start = Instant::now();
                    let mut n = 0;
                    while start.elapsed() < Duration::from_millis(100) {
                        std::hint::black_box(vm.invoke(f, &args).unwrap());
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
            rows.push(serde_json::json!({"pieces":pieces,"input_bytes":input.len(),"limit":limit,"median_speedup":ratios[3],"seconds_per_call":samples}));
        }
    }
    let result = serde_json::json!({"profile":"opt0 assertions and overflow checks","pairs":7,"minimum_sample_ms":100,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../SPLIT_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
