use vybe_compiler::primitives::{bundle, callable, dict, globals, platforms, sets};
use vybe_runtime::{
    Chunk, Op, VM, Value,
    value::{Object, ObjectKind},
};
#[path = "fixtures/collection_private_before.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
#[path = "fixtures/collection_private_raw_null.rs"]
mod original_raw;
fn output(n: &str, v: &serde_json::Value) {
    let p = std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap());
    std::fs::write(
        p.join(format!("COLLECTION_PRIVATE_{n}.json")),
        serde_json::to_vec_pretty(v).unwrap(),
    )
    .unwrap();
}
fn program(fast: bool, kind: usize, repeats: usize, prefix: bool) -> Vec<Chunk> {
    let argc = if kind == 2 { 3 } else { 2 };
    let mut root = Chunk::new("private-slot-owner");
    root.emit_op_u16(Op::REF_FUNC, 1, 0);
    root.emit(0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("entry");
    c.arity = argc;
    c.param_count = argc;
    c.local_count = argc as u16;
    if prefix {
        c.emit_i32_const(23, 0);
    }
    let mut cs = vec![root, c];
    if kind == 2 {
        let helper = if fast {
            dict::build_setdefault(&mut cs[0])
        } else {
            before::build_setdefault(&mut cs[0])
        };
        cs.push(helper);
    }
    for i in 0..repeats {
        if kind == 2 {
            cs[1].emit_op_u16(Op::REF_FUNC, 2, 0);
            cs[1].emit(0, 0);
        }
        for slot in 0..argc {
            cs[1].emit_op_u16(Op::LOCAL_GET, slot as u16, 0);
        }
        match (fast, kind) {
            (false, 0) => before::emit_add_changed(&mut cs, 1, 0),
            (true, 0) => sets::emit_add_changed(&mut cs, 1, 0),
            (false, 1) => before::emit_method_has(&mut cs, 1, 0),
            (true, 1) => dict::emit_method_has(&mut cs, 1, 0),
            (_, 2) => callable::emit_direct_invoke_chunk(&mut cs[1], argc, 0),
            _ => unreachable!(),
        }
        if i + 1 < repeats || prefix {
            cs[1].emit_op(Op::DROP, 0);
        }
    }
    cs[1].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut cs);
    globals::declare_free_globals(&mut cs);
    globals::normalize_global_table(&mut cs);
    cs
}
fn worker(cs: Vec<Chunk>) -> (VM, Value) {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    let f = vm.run(cs).unwrap();
    (vm, f)
}
fn native(cs: &[Chunk]) -> Vec<u8> {
    let b = vybe_platform_wasm::write_wasm_with_options(
        cs,
        vybe_platform_wasm::WasmWriteOptions {
            include_name_section: false,
            include_vybe_metadata: false,
        },
    );
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&b)
        .unwrap();
    b
}
fn arr(v: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(v)))
}
fn set(v: Vec<Value>) -> Value {
    let mut o = Object::new();
    o.kind = ObjectKind::Set(v.into_iter().collect());
    Value::Object(vybe_runtime::heap::alloc(o))
}
fn dictionary(map: bool, tracked: usize, existing: Option<Value>) -> Value {
    let mut o = Object::new();
    if map {
        o.kind = ObjectKind::Map(Default::default());
    }
    if let Some(value) = existing {
        if let ObjectKind::Map(m) = &mut o.kind {
            m.insert(Value::String("present".into()), value);
        } else {
            o.properties.insert("present".into(), value);
        }
    }
    if tracked > 0 {
        o.properties.insert(
            "__keys".into(),
            match tracked {
                1 => arr(vec![]),
                2 => arr(vec![Value::String("present".into())]),
                3 => arr(vec![
                    Value::String("missing".into()),
                    Value::String("present".into()),
                ]),
                _ => Value::Null,
            },
        );
    }
    Value::Object(vybe_runtime::heap::alloc(o))
}
fn id(v: &Value) -> String {
    match v {
        Value::F64(x) => format!("f64:{:x}", x.to_bits()),
        Value::Object(o) => format!("object:{:p}", std::sync::Arc::as_ptr(o)),
        _ => format!("{v:?}"),
    }
}
fn signature(v: &Value) -> Vec<String> {
    let Value::Object(o) = v else {
        return vec![id(v)];
    };
    let o = o.lock().unwrap();
    match &o.kind {
        ObjectKind::Set(s) => s.iter().map(id).collect(),
        ObjectKind::Map(m) => m
            .iter()
            .map(|(k, v)| format!("{}={}", id(k), id(v)))
            .collect(),
        _ => o
            .properties
            .iter()
            .filter(|(k, _)| k.as_str() != "__keys")
            .map(|(k, v)| format!("{k}={}", id(v)))
            .collect(),
    }
}
#[test]
fn set_membership_dictionary_values_mutations_identity_and_errors() {
    let payload = arr(vec![Value::I32(9)]);
    let values = vec![
        Value::Null,
        Value::Undefined,
        Value::Bool(false),
        Value::I32(0),
        Value::I32(7),
        Value::F64(-0.0),
        Value::F64(f64::NAN),
        Value::String("".into()),
        Value::String("é😀".into()),
        payload.clone(),
    ];
    let mut checks = 0;
    let mut sizes = vec![];
    for kind in 0..3 {
        let old = program(false, kind, 1, false);
        let new = program(true, kind, 1, false);
        let n = [native(&old), native(&new)];
        native_export::export_chunks(&old, &format!("collection-private-{kind}-before"));
        native_export::export_chunks(&new, &format!("collection-private-{kind}-after"));
        let body = if kind == 2 { 2 } else { 1 };
        assert!(new[body].code.len() < old[body].code.len());
        assert!(new[body].local_count < old[body].local_count);
        sizes.push(serde_json::json!({"kind":kind,"old_body":old[body].code.len(),"new_body":new[body].code.len(),"old_locals":old[body].local_count,"new_locals":new[body].local_count,"old_native":n[0].len(),"new_native":n[1].len()}));
        let (va, fa) = worker(old);
        let (vb, fb) = worker(new);
        let mut vs = [va, vb];
        let fs = [fa, fb];
        for existing in &values {
            for key in &values {
                if kind == 0 {
                    let mut results = vec![];
                    for side in 0..2 {
                        let input = set(vec![existing.clone()]);
                        let r = vs[side]
                            .invoke(&fs[side], &[input.clone(), key.clone()])
                            .unwrap();
                        results.push((id(&r), signature(&input)));
                        checks += 1;
                    }
                    assert_eq!(results[0], results[1]);
                }
            }
        }
        if kind > 0 {
            for map in [false, true] {
                for tracked in 0..5 {
                    for value in values.iter().map(|v| Some(v.clone())).chain([None]) {
                        for key in ["present", "missing"] {
                            let mut results = vec![];
                            for side in 0..2 {
                                let input = dictionary(map, tracked, value.clone());
                                let mut args = vec![input.clone(), Value::String(key.into())];
                                if kind == 2 {
                                    args.push(payload.clone());
                                }
                                let r = vs[side].invoke(&fs[side], &args).unwrap();
                                results.push((id(&r), signature(&input)));
                                checks += 1;
                            }
                            assert_eq!(
                                results[0], results[1],
                                "kind={kind} map={map} tracked={tracked} value={value:?} key={key}"
                            );
                        }
                    }
                }
            }
        }
        for x in [
            Value::Null,
            Value::Undefined,
            Value::I32(3),
            Value::Bool(false),
            Value::String("x".into()),
        ] {
            let mut results = vec![];
            for fast in [false, true] {
                let (mut v, f) = worker(program(fast, kind, 1, false));
                results.push(
                    v.invoke(
                        &f,
                        &[x.clone(), Value::String("missing".into()), payload.clone()],
                    )
                    .map(|x| id(&x))
                    .map_err(|e| e.message),
                );
                checks += 1;
            }
            assert_eq!(results[0], results[1], "invalid kind={kind} input={x:?}");
        }
    }
    output("GATES", &serde_json::json!({"checks":checks,"sizes":sizes}));
}
#[test]
fn lower_caller_operands_remain_live() {
    let mut checks = 0;
    for kind in 0..3 {
        for fast in [false, true] {
            let cs = program(fast, kind, 1, true);
            native(&cs);
            let (mut v, f) = worker(cs);
            let receiver = if kind == 0 {
                set(vec![])
            } else {
                dictionary(false, 2, None)
            };
            assert_eq!(
                v.invoke(
                    &f,
                    &[receiver, Value::String("present".into()), Value::I32(8)]
                )
                .unwrap(),
                Value::I32(23)
            );
            checks += 1;
        }
    }
    output("PREFIX", &serde_json::json!({"checks":checks}));
}

#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for (kind, label, tracked, present, control) in [
        (0, "set-present", 0, true, false),
        (0, "set-absent", 0, false, false),
        (1, "dict-plain-present", 0, true, false),
        (1, "dict-plain-missing", 0, false, false),
        (1, "dict-tracked-present", 2, true, false),
        (1, "dict-tracked-missing", 2, false, false),
        (2, "default-present", 0, true, false),
        (2, "default-missing", 0, false, false),
        (1, "control", 0, true, true),
    ] {
        let (va, fa) = worker(program(false, kind, 1, false));
        let (vb, fb) = worker(program(!control, kind, 1, false));
        let mut vs = [va, vb];
        let fs = [fa, fb];
        let inputs = if kind == 0 {
            [set(vec![Value::I32(7)]), set(vec![Value::I32(7)])]
        } else {
            [
                dictionary(
                    false,
                    tracked,
                    if present { Some(Value::I32(7)) } else { None },
                ),
                dictionary(
                    false,
                    tracked,
                    if present { Some(Value::I32(7)) } else { None },
                ),
            ]
        };
        let key = if kind == 0 {
            Value::I32(if present { 7 } else { 9 })
        } else {
            Value::String("present".into())
        };
        let args = [
            vec![inputs[0].clone(), key.clone(), Value::I32(91)],
            vec![inputs[1].clone(), key, Value::I32(91)],
        ];
        let mut wall = [vec![], vec![]];
        let mut cpu = [vec![], vec![]];
        let mut calls = vec![];
        for pair in 0..7 {
            let mut secs = [0., 0.];
            let mut cpus = [0., 0.];
            let mut counts = [0usize, 0];
            let mut turn = pair;
            while secs[0] < 0.1 || secs[1] < 0.1 || counts.iter().any(|n| *n < 8) {
                for side in if turn % 2 == 0 { [0, 1] } else { [1, 0] } {
                    if secs[side] >= 0.1 && counts[side] >= 8 {
                        continue;
                    }
                    // Restore the input outside EVERY timed mutator call.
                    let Value::Object(o) = &inputs[side] else {
                        unreachable!()
                    };
                    {
                        let mut o = o.lock().unwrap();
                        if kind == 0 {
                            let ObjectKind::Set(s) = &mut o.kind else {
                                unreachable!()
                            };
                            s.clear();
                            s.insert(Value::I32(7));
                        } else if present {
                            o.properties.insert("present".into(), Value::I32(7));
                        } else {
                            o.properties.remove("present");
                        }
                    }
                    let c = thread_cpu_seconds();
                    let t = Instant::now();
                    std::hint::black_box(vs[side].invoke(&fs[side], &args[side]).unwrap());
                    secs[side] += t.elapsed().as_secs_f64();
                    cpus[side] += thread_cpu_seconds() - c;
                    counts[side] += 1;
                }
                turn += 1;
            }
            for side in 0..2 {
                wall[side].push(secs[side] / counts[side] as f64);
                cpu[side].push(cpus[side] / counts[side] as f64);
            }
            calls.push(counts);
        }
        let mut wr: Vec<_> = wall[0].iter().zip(&wall[1]).map(|(a, b)| a / b).collect();
        wr.sort_by(f64::total_cmp);
        let mut cr: Vec<_> = cpu[0].iter().zip(&cpu[1]).map(|(a, b)| a / b).collect();
        cr.sort_by(f64::total_cmp);
        rows.push(serde_json::json!({"kind":kind,"label":label,"median_speedup":wr[3],"median_thread_cpu_speedup":cr[3],"wall":wall,"cpu":cpu,"calls":calls}));
    }
    output(
        "MEASUREMENTS",
        &serde_json::json!({"pairs":7,"minimum_ms_per_side":100,"separate_vms":true,"cases":rows}),
    );
}
#[cfg(target_os = "macos")]
fn thread_cpu_seconds() -> f64 {
    #[repr(C)]
    struct Timespec {
        sec: i64,
        nsec: i64,
    }
    unsafe extern "C" {
        fn clock_gettime(clock: i32, result: *mut Timespec) -> i32;
    }
    let mut t = Timespec { sec: 0, nsec: 0 };
    assert_eq!(unsafe { clock_gettime(16, &mut t) }, 0);
    t.sec as f64 + t.nsec as f64 * 1e-9
}
#[cfg(not(target_os = "macos"))]
fn thread_cpu_seconds() -> f64 {
    0.0
}

#[test]
fn missing_defaults_are_repaired_in_the_primitive() {
    let payload = arr(vec![Value::I32(9)]);
    for map in [false, true] {
        for existing in [None, Some(Value::Undefined)] {
            let input = dictionary(map, 0, existing.clone());
            let mut cs = program(false, 2, 1, false);
            cs[2] = original_raw::build_setdefault(&mut Chunk::new("unused"));
            native(&cs);
            let (mut old, f) = worker(cs);
            assert_eq!(
                old.invoke(
                    &f,
                    &[input, Value::String("present".into()), payload.clone()]
                )
                .unwrap(),
                Value::Undefined
            );
            for fast in [false, true] {
                let input = dictionary(map, 0, existing.clone());
                let (mut v, f) = worker(program(fast, 2, 1, false));
                let result = v
                    .invoke(
                        &f,
                        &[input, Value::String("present".into()), payload.clone()],
                    )
                    .unwrap();
                assert_eq!(result, payload);
            }
        }
    }
}

#[test]
fn export_original_native_regressions() {
    for kind in [1, 2] {
        let argc = if kind == 2 { 3 } else { 2 };
        let mut root = Chunk::new("original-private-owner");
        root.emit_op_u16(Op::REF_FUNC, 1, 0);
        root.emit(0, 0);
        root.emit_op(Op::RETURN, 0);
        let mut entry = Chunk::new("entry");
        entry.arity = argc;
        entry.param_count = argc;
        entry.local_count = argc as u16;
        let mut cs = vec![root, entry];
        if kind == 2 {
            cs.push(original_raw::build_setdefault(&mut Chunk::new("unused")));
            cs[1].emit_op_u16(Op::REF_FUNC, 2, 0);
            cs[1].emit(0, 0);
        }
        for slot in 0..argc {
            cs[1].emit_op_u16(Op::LOCAL_GET, slot as u16, 0);
        }
        if kind == 1 {
            original_raw::emit_method_has(&mut cs, 1, 0);
        } else {
            callable::emit_direct_invoke_chunk(&mut cs[1], argc, 0);
        }
        cs[1].emit_op(Op::RETURN, 0);
        bundle::finalize_with_runtime_helpers(&mut cs);
        globals::declare_free_globals(&mut cs);
        globals::normalize_global_table(&mut cs);
        native(&cs);
        native_export::export_chunks(&cs, &format!("collection-private-{kind}-original"));
    }
}
