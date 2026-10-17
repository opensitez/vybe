use std::sync::{Arc, Mutex};
use vybe_compiler::primitives::{self, collections};
use vybe_runtime::value::{Function, Object, ObjectKind};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/binary_search_truthiness_push.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool, cmp: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("binary-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("binary");
    c.arity = if cmp { 3 } else { 2 };
    c.param_count = c.arity;
    c.local_count = c.arity as u16;
    for slot in 0..c.arity {
        c.emit_op_u16(Op::LOCAL_GET, slot as u16, 0);
    }
    let mut cs = vec![root, c];
    match (fast, cmp) {
        (false, false) => before::emit_binary_search_pair(&mut cs, 1, 0),
        (false, true) => before::emit_binary_search_func_pair(&mut cs, 1, 0),
        (true, false) => collections::emit_binary_search_pair(&mut cs, 1, 0),
        (true, true) => collections::emit_binary_search_func_pair(&mut cs, 1, 0),
    }
    cs[1].emit_op(Op::RETURN, 0);
    if cmp {
        let mut p = Chunk::new("comparator");
        p.arity = 2;
        p.param_count = 2;
        p.local_count = 2;
        for s in 0..2 {
            p.emit_op_u16(Op::LOCAL_GET, s, 0);
        }
        let i = p.add_import("test:binary", "compare");
        p.emit_call(i, 2, 0);
        p.emit_op(Op::RETURN, 0);
        cs.push(p);
        let mut f = Chunk::new("comparator-owner");
        primitives::functions::emit_ref_func(&mut f, 2, 0, 0);
        f.emit_op(Op::RETURN, 0);
        cs.push(f);
    }
    primitives::bundle::finalize_with_runtime_helpers(&mut cs);
    primitives::globals::declare_free_globals(&mut cs);
    primitives::globals::normalize_global_table(&mut cs);
    let bindings:Vec<_>=cs[0].global_inits.iter().filter_map(|init|if let vybe_runtime::chunk::ConstExpr::RefFunc(index)=init.init{Some((init.name.clone(),index))}else{None}).collect();
    for(name,index)in bindings{let mut f=Chunk::new(format!("binding-factory:{name}"));f.emit_op_u16(Op::REF_FUNC,index as u16,0);f.emit(0,0);f.emit_op(Op::RETURN,0);cs.push(f);}
    cs
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
fn worker(
    fast: bool,
    cmp: bool,
    compare: Box<dyn Fn(&mut vybe_runtime::HostContext, &[Value]) -> Value + Send + Sync>,
) -> (VM, Value, Value) {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    vm.register_host_fn("test:binary", "compare", compare);
    let f = vm.run(program(fast, cmp)).unwrap();
    let mut p = Object::new();
    p.kind = ObjectKind::Function(Function {
        name: Some("comparator".into()),
        arity: 2,
        chunk_index: 2,
        upvalues: vec![],
    });
    (vm, f, Value::Object(vybe_runtime::heap::alloc(p)))
}
fn compare(a: &[Value]) -> Value {
    Value::F64(a[0].as_f64() - a[1].as_f64())
}
#[test]
fn binary_search_preserves_pair_and_comparator_trace() {
    let mut checks = 0;
    for cmp in [false, true] {
        for length in [0usize, 1, 2, 3, 16, 127] {
            for target in [-2i32, 0, 1, 2, 3, 15, 63, 127, 255] {
                let mut results = vec![];
                for fast in [false, true] {
                    let log = Arc::new(Mutex::new(vec![]));
                    let trace = log.clone();
                    let (mut vm, f, p) = worker(
                        fast,
                        cmp,
                        Box::new(move |_, a| {
                            trace.lock().unwrap().push((a[0].as_f64(), a[1].as_f64()));
                            compare(a)
                        }),
                    );
                    let input = array((0..length).map(|i| Value::I32((i / 2) as i32)).collect());
                    let mut args = vec![input, Value::I32(target)];
                    if cmp {
                        args.push(p);
                    }
                    let out = vm.invoke(&f, &args).unwrap();
                    let pair = items(&out);
                    let again = vm.invoke(&f, &args).unwrap();
                    let (Value::Object(o), Value::Object(q)) = (&out, &again) else {
                        panic!("array")
                    };
                    assert!(!Arc::ptr_eq(o, q));
                    results.push((pair, log.lock().unwrap().clone()));
                }
                assert_eq!(results[0], results[1]);
                let index = (0..length)
                    .find(|i| (*i / 2) as i32 >= target)
                    .unwrap_or(length);
                assert_eq!(
                    results[1].0,
                    vec![
                        Value::I32(index as i32),
                        Value::Bool(index < length && (index / 2) as i32 == target)
                    ]
                );
                checks += 1;
            }
        }
    }
    eprintln!("Verified {checks} binary search numeric/trace/allocation cases");
}
#[test]
fn natural_types_and_callback_errors_preserved() {
    let mut checks = 0;
    for values in [
        vec![],
        vec![
            Value::String("a".into()),
            Value::String("b".into()),
            Value::String("b".into()),
            Value::String("z".into()),
        ],
        vec![Value::Null, Value::Undefined],
        vec![Value::F64(f64::NAN)],
        vec![Value::Bool(false), Value::Bool(true)],
        vec![Value::bigint_i64(1), Value::bigint_i64(5)],
    ] {
        for target in [
            Value::String("b".into()),
            Value::I32(0),
            Value::Undefined,
            Value::Null,
            Value::Bool(false),
            Value::F64(f64::NAN),
            Value::bigint_i64(2),
        ] {
            let mut results = vec![];
            for fast in [false, true] {
                let (mut vm, f, _) = worker(fast, false, Box::new(|_, a| compare(a)));
                results.push(
                    vm.invoke(&f, &[array(values.clone()), target.clone()])
                        .map(|v| items(&v))
                        .map_err(|e| e.message),
                );
            }
            assert_eq!(results[0], results[1]);
            checks += 1;
        }
    }
    for answer in [
        Value::F64(f64::NAN),
        Value::String("0".into()),
        Value::Bool(false),
        Value::Undefined,
    ] {
        let mut results = vec![];
        for fast in [false, true] {
            let a = answer.clone();
            let (mut vm, f, p) = worker(fast, true, Box::new(move |_, _| a.clone()));
            results.push(
                vm.invoke(&f, &[array(vec![Value::I32(1)]), Value::I32(1), p])
                    .map(|v| items(&v))
                    .map_err(|e| e.message),
            );
        }
        assert_eq!(results[0], results[1]);
        checks += 1;
    }
    let mut errors = vec![];
    for fast in [false, true] {
        let (mut vm, f, p) = worker(
            fast,
            true,
            Box::new(|ctx, _| {
                ctx.throw_value(Value::String("compare-error".into()));
                Value::Undefined
            }),
        );
        errors.push(
            vm.invoke(&f, &[array(vec![Value::I32(1)]), Value::I32(1), p])
                .map(|v| items(&v))
                .map_err(|e| e.message),
        );
    }
    assert_eq!(errors[0], errors[1]);
    assert!(errors[1].is_err());
    eprintln!("Verified {checks} mixed-type comparator/equality cases plus throwing comparator");
}
#[test]
fn size_and_native_gates() {
    let mut rows = vec![];
    for cmp in [false, true] {
        let cs = [program(false, cmp), program(true, cmp)];
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
            if std::env::var_os("VYBE_PRIMITIVE_BINARY_EXPORT").is_some() {
                native_export::export_chunks(chunks, &format!("binary-{cmp}-{side}"));
            }
            sizes.push(serde_json::json!({"internal":chunks[1].code.len(),"native":bytes.len(),"locals":chunks[1].local_count}));
        }
        assert!(cs[1][1].code.len() < cs[0][1].code.len());
        assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
        assert!(cs[1][1].local_count < cs[0][1].local_count);
        rows.push(serde_json::json!({"comparator":cmp,"sizes":sizes}));
    }
    eprintln!("{rows:?}");
    if std::env::var_os("VYBE_PRIMITIVE_BINARY_EXPORT").is_some() {
        std::fs::write(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../BINARY_GATES.json"),
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
    for cmp in [false, true] {
        for length in [0usize, 8, 4096] {
            let mut workers =
                [false, true].map(|fast| worker(fast, cmp, Box::new(|_, a| compare(a))));
            let input = array((0..length).map(|i| Value::I32(i as i32)).collect());
            let mut samples = [vec![], vec![]];
            for pair in 0..7 {
                for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                    let (vm, f, p) = &mut workers[side];
                    let mut args = vec![input.clone(), Value::I32((length / 2) as i32)];
                    if cmp {
                        args.push(p.clone());
                    }
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
            rows.push(serde_json::json!({"comparator":cmp,"length":length,"median_speedup":ratios[3],"seconds_per_call":samples}));
        }
    }
    let result = serde_json::json!({"profile":"opt0 assertions and overflow checks enabled","pairs":7,"minimum_sample_ms":100,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../BINARY_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
