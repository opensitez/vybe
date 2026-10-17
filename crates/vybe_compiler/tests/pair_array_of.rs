use vybe_compiler::primitives::{bundle, collections, globals, platforms};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/pair_push_builder.rs"]
mod before;
fn array(v: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(v)))
}
fn items(v: &Value) -> Vec<Value> {
    let Value::Object(o) = v else { panic!("array") };
    let o = o.lock().unwrap();
    let ObjectKind::Array(v) = &o.kind else {
        panic!("array")
    };
    v.clone()
}
pub(crate) fn program(fast: bool, op: u8) -> Vec<Chunk> {
    let mut root = Chunk::new("pair-owner");
    root.emit_op_u16(Op::REF_FUNC, 1, 0);
    root.emit(0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("pair");
    if op == 2 {
        c = if fast {
            collections::build_enumerate(&mut root)
        } else {
            before::build_enumerate(&mut root)
        };
        c.name = "pair".into();
    } else {
        c.arity = 2;
        c.param_count = 2;
        c.local_count = 2;
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_op_u16(Op::LOCAL_GET, 1, 0);
        if op == 0 {
            let mut code = vec![c];
            if fast {
                collections::emit_array_pair(&mut code, 0, 0);
            } else {
                before::emit_array_pair(&mut code, 0, 0);
            }
            c = code.remove(0);
        } else if fast {
            collections::emit_array_pair_into(&mut root, &mut c, 0);
        } else {
            before::emit_array_pair_into(&mut root, &mut c, 0);
        }
        c.emit_op(Op::RETURN, 0);
    }
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
fn pair_order_values_identity_and_allocation_are_preserved() {
    let object = array(vec![Value::I32(9)]);
    for op in [0, 1] {
        for fast in [false, true] {
            let (mut vm, f) = worker(fast, op);
            for args in [
                [Value::I32(7), Value::String("x".into())],
                [Value::Null, Value::Undefined],
                [object.clone(), Value::Bool(true)],
            ] {
                let a = vm.invoke(&f, &args).unwrap();
                let b = vm.invoke(&f, &args).unwrap();
                assert_eq!(items(&a), args);
                let (Value::Object(a), Value::Object(b)) = (&a, &b) else {
                    panic!("arrays")
                };
                assert!(!std::sync::Arc::ptr_eq(a, b));
                if let Value::Object(expected) = &args[0] {
                    let Value::Object(actual) = &items(&Value::Object(a.clone()))[0] else {
                        panic!("element")
                    };
                    assert!(std::sync::Arc::ptr_eq(actual, expected));
                }
            }
        }
    }
}
#[test]
fn enumerate_matches_existing_values() {
    for input in [
        array(vec![]),
        array(vec![
            Value::Null,
            Value::Bool(true),
            Value::String("λ".into()),
        ]),
        Value::String("abc".into()),
    ] {
        let mut outcomes = vec![];
        for fast in [false, true] {
            let (mut vm, f) = worker(fast, 2);
            outcomes.push(
                items(&vm.invoke(&f, &[input.clone()]).unwrap())
                    .iter()
                    .map(items)
                    .collect::<Vec<_>>(),
            );
        }
        assert_eq!(outcomes[0], outcomes[1]);
        for (i, pair) in outcomes[1].iter().enumerate() {
            assert_eq!(pair[0], Value::I32(i as i32));
        }
    }
}
#[test]
fn native_size_and_validation() {
    let mut rows = vec![];
    for op in [0, 1, 2] {
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
        assert!(sizes[1].2 + 2 <= sizes[0].2);
        rows.push(serde_json::json!({"operation":op,"before":sizes[0],"after":sizes[1]}));
    }
    if let Ok(path) = std::env::var("PAIR_GATES_PATH") {
        std::fs::write(path, serde_json::to_vec_pretty(&rows).unwrap()).unwrap();
    }
}
fn benchmark(op: u8, path: &str) {
    let args = if op == 2 {
        vec![array((0..4096).map(Value::I32).collect())]
    } else {
        vec![Value::I32(7), Value::String("x".into())]
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
    if let Ok(path) = std::env::var(path) {
        std::fs::write(path,serde_json::to_vec_pretty(&serde_json::json!({"median":sorted[3],"ratios":ratios,"raw":raw,"scope":if op==2{"opt0 enumerate of 4096 array elements, VM invocation and full materialization"}else{"opt0 two-element array pair, VM invocation and full allocation"}})).unwrap()).unwrap();
    }
}
#[test]
#[ignore = "Seven paired opt0 pair construction batches"]
fn debug_pairs() {
    benchmark(0, "PAIR_MEASUREMENTS_PATH")
}
#[test]
#[ignore = "Seven paired opt0 enumerate materialization batches"]
fn debug_enumerate() {
    benchmark(2, "ENUMERATE_MEASUREMENTS_PATH")
}
