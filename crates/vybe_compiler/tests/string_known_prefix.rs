use vybe_compiler::primitives::{bundle, globals, platforms, strings};
use vybe_runtime::value::Object;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/string_known_prefix.rs"]
mod before;
pub(crate) fn program(fast: bool, op: u8) -> Vec<Chunk> {
    let mut root = Chunk::new("render-owner");
    root.emit_op_u16(Op::REF_FUNC, 1, 0);
    root.emit(0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = match (fast, op) {
        (false, 0) => before::build_to_string(&mut root),
        (true, 0) => strings::build_to_string(&mut root),
        (false, 1) => before::build_str_insert(&mut root),
        (true, 1) => strings::build_str_insert(&mut root),
        (false, 2) => before::build_pascal_str_insert(&mut root),
        (true, 2) => strings::build_pascal_str_insert(&mut root),
        (false, 3) => before::build_pascal_str_remove_range(&mut root),
        (true, 3) => strings::build_pascal_str_remove_range(&mut root),
        _ => unreachable!(),
    };
    c.name = "render".into();
    let mut code = vec![root, c];
    bundle::finalize_with_runtime_helpers(&mut code);
    globals::declare_free_globals(&mut code);
    globals::normalize_global_table(&mut code);
    code
}
fn worker(fast: bool, op: u8) -> (VM, Value) {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    let f = vm.run(program(fast, op)).unwrap();
    (vm, f)
}
#[test]
fn known_string_paths_preserve_values_and_strict_errors() {
    let vals = [
        Value::I32(3),
        Value::F64(-0.0),
        Value::F64(f64::NAN),
        Value::F64(f64::INFINITY),
        Value::Null,
        Value::Undefined,
        Value::Bool(false),
        Value::Bool(true),
        Value::String("".into()),
        Value::String("λ".into()),
        Value::bigint_i64(2),
        Value::Object(vybe_runtime::heap::alloc(Object::new())),
    ];
    for op in [0, 1, 2] {
        for v in &vals {
            let args = match op {
                0 => vec![v.clone()],
                1 => vec![Value::String("abcd".into()), Value::I32(2), v.clone()],
                _ => vec![v.clone(), Value::String("abcd".into()), Value::I32(2)],
            };
            let mut out = vec![];
            for fast in [false, true] {
                let (mut vm, f) = worker(fast, op);
                out.push(vm.invoke(&f, &args).map_err(|e| e.message));
            }
            assert_eq!(out[0], out[1], "op={op}, {args:?}");
        }
    }
    for start in [
        Value::I32(-1),
        Value::I32(0),
        Value::I32(2),
        Value::F64(1.5),
    ] {
        for count in [
            Value::I32(0),
            Value::I32(2),
            Value::Null,
            Value::String("2".into()),
        ] {
            let args = [Value::String("abcdef".into()), start.clone(), count];
            let mut out = vec![];
            for fast in [false, true] {
                let (mut vm, f) = worker(fast, 3);
                out.push(vm.invoke(&f, &args).map_err(|e| e.message));
            }
            assert_eq!(out[0], out[1]);
        }
    }
}
#[test]
fn size_and_native_validation() {
    let mut rows = vec![];
    for op in [0, 1, 2, 3] {
        let mut sizes = vec![];
        for fast in [false, true] {
            let code = program(fast, op);
            let wasm = vybe_platform_wasm::write_wasm(&code);
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&wasm)
                .unwrap();
            sizes.push((code[1].code.len(), wasm.len(), code[1].local_count));
        }
        assert!(sizes[1].0 < sizes[0].0);
        assert!(sizes[1].1 < sizes[0].1);
        rows.push(serde_json::json!({"operation":op,"before":sizes[0],"after":sizes[1]}));
    }
    if let Ok(path) = std::env::var("RENDER_GATES_PATH") {
        std::fs::write(path, serde_json::to_vec_pretty(&rows).unwrap()).unwrap();
    }
}
#[test]
#[ignore = "Seven paired opt0 known-string builders"]
fn debug_pairs() {
    let mut rows = vec![];
    for op in [0, 1, 2, 3] {
        let args = match op {
            0 => vec![Value::F64(12345.5)],
            1 => vec![
                Value::String("abcd".into()),
                Value::I32(2),
                Value::F64(12345.5),
            ],
            2 => vec![
                Value::F64(12345.5),
                Value::String("abcd".into()),
                Value::I32(2),
            ],
            _ => vec![Value::String("abcdef".into()), Value::I32(2), Value::I32(2)],
        };
        let mut workers = [worker(false, op), worker(true, op)];
        let mut counts = [1, 1];
        let mut ratios = vec![];
        let mut raw = vec![];
        for pair in 0..7 {
            let mut ns = [0, 0];
            for side in [pair % 2, 1 - pair % 2] {
                loop {
                    let start = std::time::Instant::now();
                    for _ in 0..counts[side] {
                        let w = &mut workers[side];
                        std::hint::black_box(w.0.invoke(&w.1, &args).unwrap());
                    }
                    ns[side] = start.elapsed().as_nanos();
                    if ns[side] >= 100_000_000 {
                        break;
                    }
                    counts[side] *= 2;
                }
            }
            ratios.push((ns[0] as f64 / counts[0] as f64) / (ns[1] as f64 / counts[1] as f64));
            raw.push(serde_json::json!({"counts":counts,"ns":ns}));
        }
        let mut sorted = ratios.clone();
        sorted.sort_by(f64::total_cmp);
        rows.push(serde_json::json!({"operation":op,"median":sorted[3],"ratios":ratios,"raw":raw}));
    }
    if let Ok(path) = std::env::var("RENDER_MEASUREMENTS_PATH") {
        std::fs::write(path, serde_json::to_vec_pretty(&rows).unwrap()).unwrap();
    }
}
