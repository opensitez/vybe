use std::sync::{Arc, Mutex};
use vybe_compiler::primitives::{self, collections};
use vybe_runtime::value::{Function, Object, ObjectKind};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/map_delete_dynamic_bound.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool) -> Vec<Chunk> {
    let mut owner = Chunk::new("map-delete-owner");
    primitives::functions::emit_ref_func(&mut owner, 1, 0, 0);
    owner.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("map-delete");
    c.arity = 2;
    c.param_count = 2;
    c.local_count = 2;
    for slot in 0..2 {
        c.emit_op_u16(Op::LOCAL_GET, slot, 0);
    }
    let mut cs = vec![owner, c];
    if fast {
        collections::emit_map_delete_func(&mut cs, 1, 0);
    } else {
        before::emit_map_delete_func(&mut cs, 1, 0);
    }
    cs[1].emit_op(Op::RETURN, 0);
    let mut p = Chunk::new("predicate");
    p.arity = 2;
    p.param_count = 2;
    p.local_count = 2;
    for slot in 0..2 {
        p.emit_op_u16(Op::LOCAL_GET, slot, 0);
    }
    let i = p.add_import("test:map-delete", "predicate");
    p.emit_call(i, 2, 0);
    p.emit_op(Op::RETURN, 0);
    cs.push(p);
    let mut factory = Chunk::new("predicate-owner");
    primitives::functions::emit_ref_func(&mut factory, 2, 0, 0);
    factory.emit_op(Op::RETURN, 0);
    cs.push(factory);
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
fn worker(
    fast: bool,
    predicate: Box<dyn Fn(&mut vybe_runtime::HostContext, &[Value]) -> Value + Send + Sync>,
) -> (VM, Value, Value) {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    vm.register_host_fn("test:map-delete", "predicate", predicate);
    let f = vm.run(program(fast)).unwrap();
    let mut p = Object::new();
    p.kind = ObjectKind::Function(Function {
        name: Some("predicate".into()),
        arity: 2,
        chunk_index: 2,
        upvalues: vec![],
    });
    (vm, f, Value::Object(vybe_runtime::heap::alloc(p)))
}
fn map(n: usize) -> Value {
    let mut o = Object::new();
    for i in 0..n {
        o.properties.insert(format!("k{i}"), Value::I32(i as i32));
    }
    Value::Object(vybe_runtime::heap::alloc(o))
}
fn snapshot(v: &Value) -> Vec<(String, Value)> {
    let Value::Object(o) = v else {
        panic!("object")
    };
    o.lock()
        .unwrap()
        .properties
        .iter()
        .map(|(k, v)| (k.clone(), v.clone()))
        .collect()
}
#[test]
fn map_delete_keeps_callback_order_snapshot_mutations_and_errors() {
    let mut checks = 0;
    for length in [0usize, 1, 2, 8, 128] {
        for mode in 0..5 {
            let mut outcomes = vec![];
            for fast in [false, true] {
                let input = map(length);
                let captured = input.clone();
                let trace = Arc::new(Mutex::new(vec![]));
                let log = trace.clone();
                let (mut vm, f, p) = worker(
                    fast,
                    Box::new(move |ctx, a| {
                        log.lock()
                            .unwrap()
                            .push((a[0].as_str().to_string(), a[1].as_f64()));
                        if mode == 3 && a[1].as_f64() == 0.0 {
                            let Value::Object(o) = &captured else {
                                unreachable!()
                            };
                            o.lock()
                                .unwrap()
                                .properties
                                .insert("new".into(), Value::I32(999));
                        }
                        if mode == 4 && a[1].as_f64() == 1.0 {
                            ctx.throw_value(Value::String("predicate-error".into()));
                        }
                        Value::Bool(match mode {
                            0 => false,
                            1 => true,
                            _ => a[1].as_f64() as i32 % 2 == 0,
                        })
                    }),
                );
                let result = vm.invoke(&f, &[input.clone(), p]).map_err(|e| e.message);
                outcomes.push((result, snapshot(&input), trace.lock().unwrap().clone()));
            }
            assert_eq!(outcomes[0], outcomes[1]);
            checks += 1;
        }
    }
    eprintln!("Verified {checks} map snapshot/callback/mutation/error cases");
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
        if std::env::var_os("VYBE_PRIMITIVE_MAP_DELETE_EXPORT").is_some() {
            native_export::export_chunks(chunks, &format!("map-delete-{side}"));
        }
        sizes.push(serde_json::json!({"body":chunks[1].code.len(),"all_internal":chunks.iter().map(|c|c.code.len()).sum::<usize>(),"native":bytes.len(),"locals":chunks[1].local_count,"chunks":chunks.len()}));
    }
    assert!(cs[1][1].code.len() < cs[0][1].code.len());
    assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
    eprintln!("{sizes:?}");
    if std::env::var_os("VYBE_PRIMITIVE_MAP_DELETE_EXPORT").is_some() {
        std::fs::write(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../MAP_DELETE_GATES.json"),
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
    for length in [0usize, 8, 4096] {
        let mut workers =
            [false, true].map(|fast| worker(fast, Box::new(|_, _| Value::Bool(false))));
        let input = map(length);
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let (vm, f, p) = &mut workers[side];
                let args = [input.clone(), p.clone()];
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
        rows.push(serde_json::json!({"length":length,"median_speedup":ratios[3],"seconds_per_call":samples}));
    }
    let result = serde_json::json!({"profile":"opt0 assertions and overflow checks","pairs":7,"minimum_sample_ms":100,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../MAP_DELETE_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}

#[test]
fn typed_map_keys_keep_identity_order_and_deletion() {
 let rows=vec![(Value::I32(1),Value::I32(0)),(Value::String("1".into()),Value::I32(1)),(Value::Bool(true),Value::I32(2)),(Value::bigint_i64(2),Value::I32(3))];
 let mut outcomes=vec![];
 for fast in [false,true] {
  let mut object=Object::new();object.kind=ObjectKind::Map(rows.iter().cloned().collect());
  let input=Value::Object(vybe_runtime::heap::alloc(object));
  let trace=Arc::new(Mutex::new(vec![]));let log=trace.clone();
  let(mut vm,f,p)=worker(fast,Box::new(move|_,a|{log.lock().unwrap().push((a[0].clone(),a[1].clone()));Value::Bool(a[1].as_f64()as i32%2==0)}));
  assert_eq!(vm.invoke(&f,&[input.clone(),p]).unwrap(),Value::Null);
  let Value::Object(o)=input else{unreachable!()};let object=o.lock().unwrap();let ObjectKind::Map(map)=&object.kind else{unreachable!()};
  let remaining=map.iter().map(|(k,v)|(k.clone(),v.clone())).collect::<Vec<_>>();
  assert_eq!(remaining,vec![rows[1].clone(),rows[3].clone()]);
  outcomes.push((remaining,trace.lock().unwrap().clone()));
 }
 assert_eq!(outcomes[0],outcomes[1]);assert_eq!(outcomes[1].1,rows);
}
