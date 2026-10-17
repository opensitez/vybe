use vybe_compiler::primitives::{bundle, collections, globals, ops, platforms};
use vybe_runtime::value::Object;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/reduction_add_ladders.rs"]
mod before;
pub(crate) fn program(fast: bool, sum: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("add-owner");
    root.emit_op_u16(Op::REF_FUNC, 1, 0);
    root.emit(0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = if sum {
        if fast {
            collections::build_sum(&mut root)
        } else {
            before::build_sum(&mut root)
        }
    } else {
        let mut c = Chunk::new("add");
        c.arity = 2;
        c.param_count = 2;
        c.local_count = 2;
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_op_u16(Op::LOCAL_GET, 1, 0);
        if fast {
            ops::emit_dyn_add_into(&mut root, &mut c, 0);
        } else {
            before::emit_dyn_add_into(&mut root, &mut c, 0);
        }
        c.emit_op(Op::RETURN, 0);
        c
    };
    c.name = "add".into();
    let mut code = vec![root, c];
    bundle::finalize_with_runtime_helpers(&mut code);
    globals::declare_free_globals(&mut code);
    globals::normalize_global_table(&mut code);
    code
}
fn worker(fast: bool, sum: bool) -> (VM, Value) {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    let f = vm.run(program(fast, sum)).unwrap();
    (vm, f)
}
fn equivalent(a: Result<Value, String>, b: Result<Value, String>) {
    match (&a, &b) {
        (Ok(Value::F64(a)), Ok(Value::F64(b))) if a.is_nan() && b.is_nan() => {}
        _ => assert_eq!(a, b),
    }
}
#[test]
fn operands_keep_values_and_errors() {
    let cases = [
        Value::I32(0),
        Value::I32(-3),
        Value::bigint_i64(2),
        Value::F64(-0.0),
        Value::F64(1.25),
        Value::F64(f64::NAN),
        Value::F64(f64::INFINITY),
        Value::Null,
        Value::Undefined,
        Value::Bool(false),
        Value::Bool(true),
        Value::String("".into()),
        Value::String("x".into()),
        Value::String("2".into()),
        Value::Object(vybe_runtime::heap::alloc(Object::new())),
    ];
    let mut workers = [worker(false, false), worker(true, false)];
    for a in &cases {
        for b in &cases {
            let args = [a.clone(), b.clone()];
            workers = [worker(false, false), worker(true, false)];
            let w = &mut workers[0];
            let old = w.0.invoke(&w.1, &args).map_err(|e| e.message);
            let w = &mut workers[1];
            let new = w.0.invoke(&w.1, &args).map_err(|e| e.message);
            equivalent(old, new);
        }
    }
}
#[test]
fn sum_preserves_mixed_and_empty_inputs() {
    for values in [
        vec![],
        vec![Value::I32(1), Value::F64(2.5)],
        vec![Value::Null, Value::Bool(true), Value::Undefined],
        vec![Value::String("x".into()), Value::I32(2), Value::Bool(false)],
    ] {
        let arg = Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)));
        let mut out = vec![];
        for fast in [false, true] {
            let (mut vm, f) = worker(fast, true);
            out.push(vm.invoke(&f, &[arg.clone()]).map_err(|e| e.message));
        }
        equivalent(out.remove(0), out.remove(0));
    }
}
#[test]
fn size_and_native_validation() {
    let mut rows = vec![];
    for sum in [false, true] {
        let mut sizes = vec![];
        for fast in [false, true] {
            let code = program(fast, sum);
            let wasm = vybe_platform_wasm::write_wasm(&code);
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&wasm)
                .unwrap();
            sizes.push((code[1].code.len(), wasm.len(), code[1].local_count));
        }
        eprintln!("sum={sum} sizes={sizes:?}");
        assert!(sizes[1].0 < sizes[0].0);
        assert!(sizes[1].1 < sizes[0].1);
        rows.push(serde_json::json!({"sum":sum,"before":sizes[0],"after":sizes[1]}));
    }
    if let Ok(path) = std::env::var("ADD_GATES_PATH") {
        std::fs::write(path, serde_json::to_vec_pretty(&rows).unwrap()).unwrap();
    }
}
#[test]
#[ignore = "Seven paired opt0 numeric sum batches"]
fn debug_sum() {
    let args = [Value::Object(vybe_runtime::heap::alloc(Object::new_array(
        (0..4096).map(Value::I32).collect(),
    )))];
    let mut workers = [worker(false, true), worker(true, true)];
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
                    assert_eq!(w.0.invoke(&w.1, &args).unwrap(), Value::F64(8386560.0));
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
    if let Ok(path) = std::env::var("ADD_MEASUREMENTS_PATH") {
        std::fs::write(path,serde_json::to_vec_pretty(&serde_json::json!({"median":sorted[3],"ratios":ratios,"raw":raw,"scope":"opt0 shared sum of 4096 i32 values; VM invocation and full sequential reduction"})).unwrap()).unwrap();
    }
}
