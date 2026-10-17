use vybe_compiler::primitives::{self, collections};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/map_clone_entry_loop.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("map-clone-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("map-clone");
    c.arity = 1;
    c.param_count = 1;
    c.local_count = 1;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let mut cs = vec![root, c];
    if fast {
        collections::emit_map_clone(&mut cs, 1, 0);
    } else {
        before::emit_map_clone(&mut cs, 1, 0);
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
fn object(o: vybe_runtime::value::Object) -> Value {
    Value::Object(vybe_runtime::heap::alloc(o))
}
fn array(v: Vec<Value>) -> Value {
    object(vybe_runtime::value::Object::new_array(v))
}
fn items(v: &Value) -> Vec<Value> {
    let Value::Object(o) = v else { panic!("map") };
    let o = o.lock().unwrap();
    let vybe_runtime::value::ObjectKind::Map(m) = &o.kind else {
        panic!("map")
    };
    m.iter().flat_map(|(k, v)| [k.clone(), v.clone()]).collect()
}
fn equal(a: &[Value], b: &[Value]) {
    assert_eq!(a.len(), b.len());
    for (a, b) in a.iter().zip(b) {
        match (a, b) {
            (Value::Object(a), Value::Object(b)) => assert!(std::sync::Arc::ptr_eq(a, b)),
            (Value::F64(a), Value::F64(b)) => assert_eq!(a.to_bits(), b.to_bits()),
            _ => assert_eq!(a, b),
        }
    }
}
fn input(shape: u8, values: Vec<Value>) -> Value {
    use vybe_runtime::value::{Object, ObjectKind};
    let mut o = Object::new();
    match shape {
        0 => {
            for (i, v) in values.into_iter().enumerate() {
                o.properties.insert(format!("k{i}"), v);
            }
        }
        1 => {
            o.kind = ObjectKind::Map(Default::default());
            if let ObjectKind::Map(m) = &mut o.kind {
                for (i, v) in values.into_iter().enumerate() {
                    m.insert(Value::I32(i as i32), v);
                }
            }
        }
        2 => {
            o.kind = ObjectKind::Set(values.into_iter().collect());
        }
        3 => {
            o = Object::new_array(values);
        }
        4 => {
            o = Object::new_array(values);
            if let ObjectKind::Array(v) = &o.kind {
                let len = v.len();
                o.properties.insert(
                    "__holes".into(),
                    array((0..len).step_by(2).map(|i| Value::I32(i as i32)).collect()),
                );
            }
        }
        _ => {
            for (i, v) in values.into_iter().enumerate() {
                o.properties.insert(
                    if i % 2 == 0 {
                        format!("{i}")
                    } else {
                        format!("k{i}")
                    },
                    v,
                );
            }
            o.properties.insert(
                "__keys".into(),
                array(vec![Value::String("k1".into()), Value::String("0".into())]),
            );
        }
    }
    object(o)
}
#[test]
fn clone_preserves_shapes_order_tags_identity_and_freshness() {
    let shared = array(vec![Value::I32(1)]);
    let mut workers = [worker(false), worker(true)];
    let mut checks = 0;
    for shape in 0..6 {
        for count in [0usize, 1, 2, 8, 255, 4096] {
            let values: Vec<Value> = (0..count)
                .map(|i| match i % 8 {
                    0 => Value::F64(-0.0),
                    1 => Value::F64(f64::NAN),
                    2 => Value::Null,
                    3 => Value::Undefined,
                    4 => Value::Bool(false),
                    5 => Value::bigint_i64(i as i64),
                    6 => Value::String("é😀".into()),
                    _ => shared.clone(),
                })
                .collect();
            let src = input(shape, values);
            let out = workers
                .each_mut()
                .map(|(vm, f)| vm.invoke(f, &[src.clone()]).unwrap());
            if shape == 2 {
                let Value::Object(o) = &src else {
                    unreachable!()
                };
                let o = o.lock().unwrap();
                let vybe_runtime::value::ObjectKind::Set(set) = &o.kind else {
                    unreachable!()
                };
                let expected: Vec<_> = set.iter().flat_map(|v| [v.clone(), v.clone()]).collect();
                equal(&items(&out[1]), &expected);
                if count >= 8 {
                    // Reproduce the old setter's erroneous null-key stringification.
                    assert!(
                        items(&out[0])
                            .chunks_exact(2)
                            .any(|p| p[0] == Value::String("null".into()))
                    );
                    assert!(items(&out[1]).chunks_exact(2).any(|p| p[0] == Value::Null));
                }
            } else {
                equal(&items(&out[0]), &items(&out[1]));
            }
            for (vm, f) in &mut workers {
                let a = vm.invoke(f, &[src.clone()]).unwrap();
                let b = vm.invoke(f, &[src.clone()]).unwrap();
                let (Value::Object(a), Value::Object(b)) = (a, b) else {
                    panic!("array")
                };
                assert!(!std::sync::Arc::ptr_eq(&a, &b));
            }
            checks += 1;
        }
    }
    for scalar in [
        Value::Null,
        Value::Undefined,
        Value::Bool(true),
        Value::I32(17),
        Value::String("é😀".into()),
        Value::bigint_i64(2),
    ] {
        let out = workers.each_mut().map(|(vm, f)| {
            vm.invoke(f, &[scalar.clone()])
                .map(|v| {
                    if matches!(v, Value::Null) {
                        None
                    } else {
                        Some(items(&v))
                    }
                })
                .map_err(|e| e.message)
        });
        assert_eq!(out[0], out[1]);
        checks += 1;
    }
    eprintln!("Verified {checks} map clone shape/scalar/freshness cases");
}
fn host(vm: &VM, module: &str, name: &str) -> Value {
    let mut o = vybe_runtime::value::Object::new();
    o.kind = vybe_runtime::value::ObjectKind::HostFunction(
        vm.resolve_host_function_index(module, name).unwrap(),
    );
    object(o)
}
#[test]
fn clone_preserves_getters_mutations_and_exceptions() {
    use std::sync::{Arc, Mutex};
    use vybe_runtime::value::Object;
    for mode in ["plain", "change", "delete", "append", "throw"] {
        let mut outcomes = vec![];
        for fast in [false, true] {
            let (mut vm, f) = worker(fast);
            let mut o = Object::new();
            o.properties.insert("first".into(), Value::I32(0));
            o.properties.insert("later".into(), Value::I32(2));
            let src = object(o);
            let cap = src.clone();
            let trace = Arc::new(Mutex::new(vec![]));
            let log = trace.clone();
            vm.register_host_fn(
                "test:values",
                "getter",
                Box::new(move |ctx, _| {
                    log.lock().unwrap().push("first");
                    let Value::Object(o) = &cap else {
                        unreachable!()
                    };
                    let mut o = o.lock().unwrap();
                    match mode {
                        "change" => {
                            o.properties.insert("later".into(), Value::I32(99));
                        }
                        "delete" => {
                            o.properties.shift_remove("later");
                        }
                        "append" => {
                            o.properties.insert("new".into(), Value::I32(3));
                        }
                        "throw" => {
                            ctx.throw_value(Value::String("getter-error".into()));
                        }
                        _ => {}
                    }
                    Value::I32(7)
                }),
            );
            let getter = host(&vm, "test:values", "getter");
            let mut descriptor = Object::new();
            descriptor.properties.insert("get".into(), getter);
            descriptor
                .properties
                .insert("enumerable".into(), Value::Bool(true));
            descriptor
                .properties
                .insert("configurable".into(), Value::Bool(true));
            let define = host(&vm, "ecma:object", "defineProperty");
            vm.invoke(
                &define,
                &[
                    src.clone(),
                    Value::String("first".into()),
                    object(descriptor),
                ],
            )
            .unwrap();
            let result = vm
                .invoke(&f, &[src])
                .map(|v| items(&v))
                .map_err(|e| e.message);
            outcomes.push((result, trace.lock().unwrap().clone()));
        }
        assert_eq!(outcomes[0], outcomes[1], "{mode}");
        assert_eq!(outcomes[1].1, vec!["first"]);
        match mode {
            "throw" => assert!(outcomes[1].0.is_err()),
            "change" => equal(
                outcomes[1].0.as_ref().unwrap(),
                &[
                    Value::String("first".into()),
                    Value::I32(7),
                    Value::String("later".into()),
                    Value::I32(99),
                ],
            ),
            "delete" => equal(
                outcomes[1].0.as_ref().unwrap(),
                &[
                    Value::String("first".into()),
                    Value::I32(7),
                    Value::String("later".into()),
                    Value::Undefined,
                ],
            ),
            _ => equal(
                outcomes[1].0.as_ref().unwrap(),
                &[
                    Value::String("first".into()),
                    Value::I32(7),
                    Value::String("later".into()),
                    Value::I32(2),
                ],
            ),
        }
    }
    eprintln!("Verified 5 map clone getter/mutation/error cases");
}
#[test]
fn size_and_native_exports() {
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
        if std::env::var_os("VYBE_PRIMITIVE_MAP_CLONE_EXPORT").is_some() {
            native_export::export_chunks(&cs, &format!("map-clone-{}", u8::from(fast)));
        }
        rows.push(serde_json::json!({"body":cs[1].code.len(),"native":bytes.len(),"locals":cs[1].local_count,"all_bodies":cs.iter().map(|c|c.code.len()).sum::<usize>(),"chunks":cs.len()}));
    }
    assert!(rows[1]["native"].as_u64() < rows[0]["native"].as_u64());
    assert!(rows[1]["body"].as_u64() < rows[0]["body"].as_u64());
    if std::env::var_os("VYBE_PRIMITIVE_MAP_CLONE_EXPORT").is_some() {
        std::fs::write(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../MAP_CLONE_GATES.json"),
            serde_json::to_vec_pretty(&rows).unwrap(),
        )
        .unwrap();
    }
    eprintln!("{rows:?}");
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::{Duration, Instant};
    let mut rows = vec![];
    for shape in [0u8, 1, 3] {
        for count in [0usize, 8, 4096] {
            let mut workers = [worker(false), worker(true)];
            let src = input(shape, (0..count).map(|i| Value::I32(i as i32)).collect());
            let mut samples = [vec![], vec![]];
            for pair in 0..7 {
                for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                    let (vm, f) = &mut workers[side];
                    let start = Instant::now();
                    let mut n = 0;
                    while start.elapsed() < Duration::from_millis(100) || n < 8 {
                        std::hint::black_box(vm.invoke(f, &[src.clone()]).unwrap());
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
            rows.push(serde_json::json!({"shape":shape,"count":count,"median_speedup":ratios[3],"seconds_per_call":samples}));
        }
    }
    let result = serde_json::json!({"profile":"opt0 assertions and overflow checks","pairs":7,"minimum_sample_ms":100,"minimum_calls_per_sample":8,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../MAP_CLONE_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
#[test]
fn clone_preserves_typed_keys_and_independent_contents() {
    use vybe_runtime::value::{Object, ObjectKind};
    let reference = array(vec![Value::I32(17)]);
    let keys = vec![
        Value::I32(1),
        Value::String("1".into()),
        Value::Bool(true),
        Value::bigint_i64(1),
        Value::F64(f64::NAN),
        Value::F64(-0.0),
        reference.clone(),
    ];
    let mut o = Object::new();
    o.kind = ObjectKind::Map(Default::default());
    if let ObjectKind::Map(m) = &mut o.kind {
        for k in &keys {
            m.insert(k.clone(), reference.clone());
        }
    }
    let src = object(o);
    let expected: Vec<_> = keys
        .into_iter()
        .flat_map(|k| [k, reference.clone()])
        .collect();
    // Reproduce the original key conversion, then require exact
    // shallow clone semantics from the production path.
    let (mut old_vm, old_f) = worker(false);
    let old = old_vm.invoke(&old_f, &[src.clone()]).unwrap();
    assert_eq!(items(&old).len(), 14);
    assert!(
        !items(&old)
            .chunks_exact(2)
            .any(|p| p[0] == Value::Bool(true))
    );
    assert!(
        items(&old)
            .chunks_exact(2)
            .any(|p| p[0] == Value::String("true".into()))
    );
    for fast in [true] {
        let (mut vm, f) = worker(fast);
        let out = vm.invoke(&f, &[src.clone()]).unwrap();
        equal(&items(&out), &expected);
        equal(&items(&src), &expected);
        let size = host(&vm, "ecma:map", "size");
        assert_eq!(vm.invoke(&size, &[out.clone()]).unwrap().as_i32(), 7);
        let set = host(&vm, "ecma:map", "set");
        vm.invoke(
            &set,
            &[out.clone(), Value::String("new".into()), Value::I32(99)],
        )
        .unwrap();
        assert_eq!(items(&out).len(), 16);
        equal(&items(&src), &expected);
        assert_eq!(vm.invoke(&f, &[Value::Null]).unwrap(), Value::Null);
        assert_eq!(vm.invoke(&f, &[Value::Undefined]).unwrap(), Value::Null);
    }
    eprintln!(
        "Verified 1 original typed-key failure and 1 fixed clone/content/mutation/null control"
    );
}
#[test]
fn map_from_entries_import_has_one_parameter() {
    let bytes = vybe_platform_wasm::write_wasm_with_options(
        &program(true),
        vybe_platform_wasm::WasmWriteOptions {
            include_name_section: false,
            include_vybe_metadata: false,
        },
    );
    let mut params = vec![];
    let mut found = false;
    for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
        match payload.unwrap() {
            wasmparser::Payload::TypeSection(types) => {
                for group in types {
                    for ty in group.unwrap().into_types() {
                        params.push(match ty.composite_type.inner {
                            wasmparser::CompositeInnerType::Func(f) => Some(f.params().len()),
                            _ => None,
                        });
                    }
                }
            }
            wasmparser::Payload::ImportSection(imports) => {
                for import in imports.into_imports() {
                    let i = import.unwrap();
                    if i.module == "ecma:map" && i.name == "fromEntries" {
                        let wasmparser::TypeRef::Func(t) = i.ty else {
                            panic!("function")
                        };
                        assert_eq!(params[t as usize], Some(1));
                        found = true;
                    }
                }
            }
            _ => {}
        }
    }
    assert!(found);
}
