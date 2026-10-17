#![cfg(feature = "primitive-benchmarks")]
use vybe_ast::*;
use vybe_runtime::{Chunk, VM, Value};
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool, repeats: usize) -> Vec<Chunk> {
    let profile = vybe_compiler::profile::parse_profile(
        "[compiler]\ncase_sensitive=true\nsupports_dynamic_import=true\n",
    )
    .unwrap();
    let compiler = vybe_compiler::Compiler::with_profile(profile);
    let mut body: Vec<_> = (1..repeats)
        .map(|_| Statement::new(StmtKind::Expr(Expression::ident("__js_import_meta"))))
        .collect();
    body.push(Statement::new(StmtKind::Return(Some(Expression::ident(
        "__js_import_meta",
    )))));
    vybe_compiler::primitives::expressions::benchmark_with_import_meta_elision(fast, || {
        compiler
            .compile(&Module {
                name: "meta-elision".into(),
                language: Lang::Unknown,
                imports: vec![],
                canon: Default::default(),
                directives: Default::default(),
                body: vec![
                    Statement::new(StmtKind::FunctionDecl {
                        name: "entry".into(),
                        params: vec![],
                        body,
                        return_type: None,
                        modifiers: Default::default(),
                        handles: vec![],
                        is_async: false,
                        is_generator: false,
                        is_sub: false,
                    }),
                    Statement::new(StmtKind::Return(Some(Expression::ident("entry")))),
                ],
            })
            .unwrap()
    })
}
fn worker(fast: bool) -> (VM, Value) {
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(program(fast, 1)).unwrap();
    (vm, f)
}
fn object() -> Value {
    Value::Object(vybe_runtime::heap::alloc(vybe_runtime::value::Object::new()))
}
#[test]
fn cached_values_initialization_and_sizes() {
    let values = [
        Value::Null,
        Value::Undefined,
        Value::I32(0),
        Value::I32(7),
        Value::Bool(false),
        Value::String("".into()),
        Value::String("cached".into()),
        Value::bigint_i64(0),
        Value::Symbol("symbol".into()),
        object(),
    ];
    let mut checks = 0;
    for value in values {
        for fast in [false, true] {
            let (mut vm, f) = worker(fast);
            vm.set_global("__js_import_meta", value.clone());
            let out = vm.invoke(&f, &[]).unwrap();
            if matches!(value, Value::Null | Value::Undefined) {
                assert!(matches!(out, Value::Object(_)));
                if let Value::Object(ref object) = out {
                    assert_eq!(object.lock().unwrap().get("url"), Value::String("".into()));
                }
                assert_eq!(vm.invoke(&f, &[]).unwrap(), out);
            } else {
                assert_eq!(out, value);
                assert_eq!(vm.invoke(&f, &[]).unwrap(), value);
            }
            checks += 1;
        }
    }
    let mut rows = vec![];
    for repeats in [1, 8] {
        let mut sizes = vec![];
        for fast in [false, true] {
            let cs = program(fast, repeats);
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
            native_export::export_chunks(&cs, &format!("import-meta-{repeats}-{fast}"));
            sizes.push(serde_json::json!({"body":cs.iter().find(|c|c.name=="entry").unwrap().code.len(),"locals":cs.iter().find(|c|c.name=="entry").unwrap().local_count,"native":bytes.len(),"linked_body":cs.iter().map(|c|c.code.len()).sum::<usize>()}));
        }
        assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
        assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
        rows.push(serde_json::json!({"repeats":repeats,"sizes":sizes}));
    }
    let result = serde_json::json!({"checks":checks,"rows":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../IMPORT_META_GATES.json"),
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
    for (label, repeats, cold, control) in [
        ("control", 1, false, true),
        ("warm1", 1, false, false),
        ("warm8", 8, false, false),
        ("null", 1, true, false),
    ] {
        let (mut vm, a) = worker(false);
        let b = if control {
            a.clone()
        } else {
            vm.run(program(true, repeats)).unwrap()
        };
        let a = if repeats == 1 {
            a
        } else {
            vm.run(program(false, repeats)).unwrap()
        };
        let fs = [a, b];
        let cached = object();
        vm.set_global("__js_import_meta", cached);
        for f in &fs {
            for _ in 0..8 {
                if cold {
                    vm.set_global("__js_import_meta", Value::Null);
                }
                std::hint::black_box(vm.invoke(f, &[]).unwrap());
            }
        }
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let t = Instant::now();
                let mut n = 0;
                while t.elapsed() < Duration::from_millis(500) || n < 8 {
                    if cold {
                        vm.set_global("__js_import_meta", Value::Null);
                    }
                    std::hint::black_box(vm.invoke(&fs[side], &[]).unwrap());
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
        rows.push(serde_json::json!({"label":label,"repeats":repeats,"cold":cold,"control":control,"median_speedup":ratios[3],"seconds_per_call":samples}));
    }
    let result = serde_json::json!({"same_vm":true,"pairs":7,"minimum_sample_ms":500,"warmups":8,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../IMPORT_META_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
