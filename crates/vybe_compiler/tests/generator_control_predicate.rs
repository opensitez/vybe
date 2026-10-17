use vybe_compiler::primitives::{self, generators, instructions::host, ops};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/generator_control_truthiness.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool, kind: u8, repeats: usize) -> Vec<Chunk> {
    let mut root = Chunk::new("generator-control-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("dispatch");
    c.arity = 1;
    c.param_count = 1;
    c.local_count = 1;
    for i in 0..repeats {
        let mut ret = |c: &mut Chunk, line| c.emit_op(Op::RETURN, line);
        match kind {
            0 => {
                if fast {
                    generators::emit_entry_control(&mut c, 0, 0, &mut ret)
                } else {
                    before::emit_entry_control(&mut c, 0, 0, &mut ret)
                }
            }
            1 => {
                c.emit_op_u16(Op::LOCAL_GET, 0, 0);
                if fast {
                    generators::emit_resume_dispatch(&mut c, 0, &mut ret);
                } else {
                    before::emit_resume_dispatch(&mut c, 0, &mut ret);
                }
                if i + 1 < repeats {
                    c.emit_op(Op::DROP, 0);
                }
            }
            2 => {
                c.emit_op_u16(Op::LOCAL_GET, 0, 0);
                c.emit_string_const("return", 0);
                host::emit(&mut c, "wasm:js-string", "equals", 2, 0);
                if !fast {
                    ops::emit_dyn_to_bool(&mut c, 0);
                }
                c.emit_if_i32(0);
                c.emit_i32_const(1, 0);
                c.emit_else(0);
                c.emit_i32_const(0, 0);
                c.emit_end(0);
                if i + 1 < repeats {
                    c.emit_op(Op::DROP, 0);
                }
            }
            _ => unreachable!(),
        }
    }
    if kind == 0 {
        c.emit_i32_const(99, 0);
    }
    c.emit_op(Op::RETURN, 0);
    let mut cs = vec![root, c];
    primitives::bundle::finalize_with_runtime_helpers(&mut cs);
    primitives::globals::declare_free_globals(&mut cs);
    primitives::globals::normalize_global_table(&mut cs);
    let bindings: Vec<_> = cs[0]
        .global_inits
        .iter()
        .filter_map(|g| {
            if let vybe_runtime::chunk::ConstExpr::RefFunc(i) = g.init {
                Some((g.name.clone(), i))
            } else {
                None
            }
        })
        .collect();
    for (name, i) in bindings {
        let mut f = Chunk::new(format!("binding-factory:{name}"));
        f.emit_op_u16(Op::REF_FUNC, i as u16, 0);
        f.emit(0, 0);
        f.emit_op(Op::RETURN, 0);
        cs.push(f);
    }
    cs
}
fn worker(fast: bool, kind: u8, repeats: usize) -> (VM, Value) {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(program(fast, kind, repeats)).unwrap();
    (vm, f)
}
fn packet(flag: Value, op: Value, value: Value) -> Value {
    let mut o = vybe_runtime::value::Object::new();
    o.properties.insert("__vybe_generator_control".into(), flag);
    o.properties.insert("op".into(), op);
    o.properties.insert("value".into(), value);
    Value::Object(vybe_runtime::heap::alloc(o))
}
#[test]
fn dispatch_preserves_control_effects_and_unproven_marker_truthiness() {
    let mut checks = 0;
    for kind in [0, 1] {
        let mut workers = [worker(false, kind, 1), worker(true, kind, 1)];
        for flag in [
            Value::Bool(true),
            Value::Bool(false),
            Value::I32(0),
            Value::I32(1),
            Value::String("".into()),
            Value::String("yes".into()),
            Value::Null,
            Value::Undefined,
        ] {
            for op in [
                Value::String("next".into()),
                Value::String("throw".into()),
                Value::String("return".into()),
                Value::String("".into()),
                Value::Null,
                Value::I32(1),
            ] {
                for value in [Value::I32(17), Value::Null, Value::String("payload".into())] {
                    let input = packet(flag.clone(), op.clone(), value.clone());
                    let out = workers
                        .each_mut()
                        .map(|(vm, f)| vm.invoke(f, &[input.clone()]).map_err(|e| e.message));
                    assert_eq!(out[0], out[1], "kind={kind} flag={flag:?} op={op:?}");
                    if let (Value::Bool(marker), Value::String(op)) = (&flag, &op) {
                        if *marker && op.as_ref() == "throw" {
                            assert!(out[1].is_err());
                        } else {
                            let expected = if *marker && op.as_ref() == "return" {
                                value
                            } else if kind == 0 {
                                Value::I32(99)
                            } else {
                                input
                            };
                            assert_eq!(out[1], Ok(expected));
                        }
                    }
                    if out[0].is_err() {
                        // An unhandled throw leaves the VM faulted. Later
                        // cases need fresh instances to test their own result.
                        workers = [worker(false, kind, 1), worker(true, kind, 1)];
                    }
                    checks += 1;
                }
            }
        }
        for input in [
            Value::Null,
            Value::Undefined,
            Value::I32(0),
            Value::Bool(true),
            Value::String("resume".into()),
        ] {
            let out = workers
                .each_mut()
                .map(|(vm, f)| vm.invoke(f, &[input.clone()]).map_err(|e| e.message));
            assert_eq!(out[0], out[1]);
            assert_eq!(out[1], Ok(if kind == 0 { Value::I32(99) } else { input }));
            checks += 1;
        }
    }
    let mut workers = [worker(false, 2, 1), worker(true, 2, 1)];
    for input in [
        Value::String("return".into()),
        Value::String("throw".into()),
        Value::String("".into()),
        Value::String("rétürn😀".into()),
        Value::Null,
        Value::Undefined,
        Value::Bool(true),
        Value::I32(1),
    ] {
        let out = workers
            .each_mut()
            .map(|(vm, f)| vm.invoke(f, &[input.clone()]).map_err(|e| e.message));
        assert_eq!(out[0], out[1]);
        if let Value::String(text) = input {
            assert_eq!(out[1], Ok(Value::I32(i32::from(text.as_ref() == "return"))));
        }
        if out[0].is_err() {
            workers = [worker(false, 2, 1), worker(true, 2, 1)];
        }
        checks += 1;
    }
    eprintln!(
        "Verified {checks} generator control/marker/value/exception and strict predicate cases"
    );
}
#[test]
fn size_and_native_exports() {
    let mut rows = vec![];
    for kind in 0..3 {
        let mut sizes = vec![];
        for fast in [false, true] {
            let cs = program(fast, kind, 1);
            let c = &cs[1];
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
            if std::env::var_os("VYBE_PRIMITIVE_GENERATOR_CONTROL_EXPORT").is_some() {
                native_export::export_chunks(
                    &cs,
                    &format!("generator-predicate-{kind}-{}", u8::from(fast)),
                );
            }
            sizes.push(serde_json::json!({"body":c.code.len(),"locals":c.local_count,"native":bytes.len()}));
        }
        assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
        rows.push(serde_json::json!({"kind":kind,"sizes":sizes}));
    }
    if std::env::var_os("VYBE_PRIMITIVE_GENERATOR_CONTROL_EXPORT").is_some() {
        std::fs::write(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../GENERATOR_PREDICATE_GATES.json"),
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
    for (label, kind, input) in [
        (
            "entry-control",
            0,
            packet(Value::Bool(true), Value::String("next".into()), Value::Null),
        ),
        (
            "resume-control",
            1,
            packet(Value::Bool(true), Value::String("next".into()), Value::Null),
        ),
        ("predicate-false", 2, Value::String("next".into())),
        ("predicate-true", 2, Value::String("return".into())),
        ("entry-scalar-control", 0, Value::I32(7)),
        ("resume-scalar-control", 1, Value::I32(7)),
    ] {
        let mut workers = [worker(false, kind, 64), worker(true, kind, 64)];
        for (vm, f) in &mut workers {
            for _ in 0..8 {
                std::hint::black_box(vm.invoke(f, &[input.clone()]).unwrap());
            }
        }
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let (vm, f) = &mut workers[side];
                let start = Instant::now();
                let mut count = 0;
                while start.elapsed() < Duration::from_millis(200) || count < 8 {
                    std::hint::black_box(vm.invoke(f, &[input.clone()]).unwrap());
                    count += 1;
                }
                samples[side].push(start.elapsed().as_secs_f64() / count as f64);
            }
        }
        let mut ratios: Vec<_> = samples[0]
            .iter()
            .zip(&samples[1])
            .map(|(a, b)| a / b)
            .collect();
        ratios.sort_by(f64::total_cmp);
        rows.push(
            serde_json::json!({"case":label,"median_speedup":ratios[3],"seconds_per_call":samples}),
        );
    }
    let result = serde_json::json!({"pairs":7,"minimum_sample_ms":200,"minimum_calls":8,"warmup_calls":8,"operations_per_call":64,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../GENERATOR_PREDICATE_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
