use std::sync::{Arc, Mutex};
use vybe_compiler::primitives::{bundle, callable, globals, platforms, polyfills, slices};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/collection_optimization_slice_predicates.rs"]
mod baseline;
fn array(values: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)))
}
fn finish(mut chunks: Vec<Chunk>) -> Vec<Chunk> {
    chunks[0].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
fn runtime(values: &[(&str, Value)]) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    for (k, v) in values {
        vm.set_global(*k, v.clone());
    }
    vm
}
fn validate(chunks: &[Chunk]) {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(chunks))
        .unwrap();
}
fn error(mut error: vybe_runtime::VMError) -> String {
    for frame in &mut error.call_stack {
        frame.offset = 0;
    }
    format!("{error:?}")
}
fn vector(value: &Value) -> Vec<Value> {
    let Value::Object(o) = value else {
        panic!("array expected")
    };
    let o = o.lock().unwrap();
    let ObjectKind::Array(v) = &o.kind else {
        panic!("array expected")
    };
    v.clone()
}
fn snapshot(value: &Value) -> String {
    let Value::Object(o) = value else {
        return format!("{value:?}");
    };
    let len = {
        let o = o.lock().unwrap();
        match &o.kind {
            ObjectKind::Array(v) => return format!("Array{v:?}"),
            ObjectKind::TypedArray(t) => t.length,
            _ => return format!("{:?}", o.kind),
        }
    };
    let mut c = Chunk::new("<typed-view-snapshot>");
    for index in 0..len {
        globals::emit_read(&mut c, "input", 0);
        c.emit_i32_const(index as i32, 0);
        let get = c.add_import("ecma:array", "get");
        c.emit_call(get, 2, 0);
    }
    c.emit_array_new_fixed(0, len as u16, 0);
    format!(
        "TypedArray{:?}",
        vector(
            &runtime(&[("input", value.clone())])
                .run(finish(vec![c]))
                .unwrap()
        )
    )
}
fn view() -> Value {
    let mut c = Chunk::new("<actual-view>");
    globals::emit_read(&mut c, "input", 0);
    let new = c.add_import("ecma:uint8array", "new");
    c.emit_call(new, 1, 0);
    runtime(&[("input", array((0..8).map(Value::I32).collect()))])
        .run(finish(vec![c]))
        .unwrap()
}
fn program(old: bool, operation: u8, raises: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("<slice-predicates>");
    if operation == 1 || operation == 4 {
        let helper = if old {
            if operation == 1 {
                baseline::build_slice_step(&mut root)
            } else {
                baseline::build_slice(&mut root)
            }
        } else {
            polyfills::build_runtime_helpers_for_exports(
                &mut root,
                &[if operation == 1 {
                    "__stdlib_slicestep"
                } else {
                    "__stdlib_slice"
                }],
            )
            .chunks
            .remove(0)
        };
        let mut chunks = vec![root, helper];
        chunks[0].emit_op_u16(Op::REF_FUNC, 1, 0);
        chunks[0].emit(0, 0);
        for name in if operation == 1 {
            vec!["input", "lower", "upper", "step"]
        } else {
            vec!["input", "lower", "upper"]
        } {
            globals::emit_read(&mut chunks[0], name, 0);
        }
        callable::emit_direct_invoke_chunk(&mut chunks[0], if operation == 1 { 4 } else { 3 }, 0);
        return finish(chunks);
    }
    let mut chunks = vec![root];
    for name in ["input", "lower", "upper"] {
        globals::emit_read(&mut chunks[0], name, 0);
    }
    if operation > 0 {
        globals::emit_read(&mut chunks[0], "step", 0);
    }
    if operation == 2 {
        globals::emit_read(&mut chunks[0], "source", 0);
    }
    if old {
        match operation {
            0 => baseline::emit_contiguous(&mut chunks, 0, 0),
            2 => baseline::emit_strided_assign(&mut chunks, 0, 0, baseline::Options::new(raises)),
            _ => baseline::emit_strided_del(&mut chunks, 0, 0, baseline::Options::new(raises)),
        }
    } else {
        match operation {
            0 => slices::emit_contiguous(&mut chunks, 0, 0),
            2 => slices::emit_strided_assign(&mut chunks, 0, 0, slices::Options::new(raises)),
            _ => slices::emit_strided_del(&mut chunks, 0, 0, slices::Options::new(raises)),
        }
    }
    if operation > 1 {
        globals::emit_read(&mut chunks[0], "input", 0);
    }
    finish(chunks)
}
#[test]
fn slicing_preserves_array_string_actual_view_and_error_outcomes() {
    for operation in [0, 1, 4] {
        let programs = [
            program(true, operation, false),
            program(false, operation, false),
        ];
        for chunks in &programs {
            validate(chunks);
        }
        for input in [
            array((0..8).map(Value::I32).collect()),
            Value::String("é😀abc".into()),
            view(),
            Value::Null,
        ] {
            for (lower, upper, step) in [
                (Value::I32(1), Value::I32(5), Value::I32(2)),
                (Value::Null, Value::Null, Value::I32(-1)),
                (Value::I32(-4), Value::I32(20), Value::I32(1)),
                (Value::String("1".into()), Value::I32(3), Value::I32(1)),
                (Value::Null, Value::Null, Value::I32(0)),
            ] {
                let outputs = programs.each_ref().map(|chunks| {
                    runtime(&[
                        ("input", input.clone()),
                        ("lower", lower.clone()),
                        ("upper", upper.clone()),
                        ("step", step.clone()),
                    ])
                    .run(chunks.clone())
                    .map(|v| snapshot(&v))
                    .map_err(error)
                });
                assert_eq!(
                    outputs[0], outputs[1],
                    "op={operation} lower={lower:?} upper={upper:?} step={step:?}"
                );
            }
        }
    }
}
#[test]
fn private_stride_counters_preserve_aliasing_directions_zero_policy_and_mutation() {
    for operation in [2, 3] {
        for raises in [false, true] {
            let programs = [
                program(true, operation, raises),
                program(false, operation, raises),
            ];
            for chunks in &programs {
                validate(chunks);
            }
            for step in [
                Value::I32(1),
                Value::I32(2),
                Value::I32(-1),
                Value::I32(-2),
                Value::F64(0.5),
                Value::I32(0),
            ] {
                for alias in [false, true] {
                    let outputs = programs.each_ref().map(|chunks| {
                        let input = array((0..8).map(Value::I32).collect());
                        let source = if alias {
                            input.clone()
                        } else {
                            array((20..36).map(Value::I32).collect())
                        };
                        let result = runtime(&[
                            ("input", input.clone()),
                            ("source", source),
                            ("lower", Value::Null),
                            ("upper", Value::Null),
                            ("step", step.clone()),
                        ])
                        .run(chunks.clone())
                        .map(|v| snapshot(&v))
                        .map_err(error);
                        (result, snapshot(&input))
                    });
                    assert_eq!(
                        outputs[0], outputs[1],
                        "op={operation} raises={raises} step={step:?} alias={alias}"
                    );
                }
            }
        }
    }
}
#[test]
fn strided_assignment_preserves_observed_get_set_sequence() {
    let outcomes = [true, false].map(|old| {
        let mut root = Chunk::new("<observed-strided-assignment>");
        let mut getitem = Chunk::new("<source-getitem>");
        getitem.arity = 2;
        getitem.local_count = 2;
        getitem.emit_string_const("source", 0);
        getitem.emit_op_u16(Op::LOCAL_GET, 1, 0);
        let observe = getitem.add_import("ecma:string", "charAt");
        getitem.emit_call(observe, 2, 0);
        getitem.emit_op(Op::RETURN, 0);
        globals::emit_read(&mut root, "source", 0);
        root.emit_string_const("__getitem__", 0);
        root.emit_op_u16(Op::REF_FUNC, 1, 0);
        root.emit(0, 0);
        let set = root.add_import("ecma:object", "set");
        root.emit_call(set, 3, 0);
        root.emit_op(Op::DROP, 0);
        for name in ["input", "lower", "upper", "step", "source"] {
            globals::emit_read(&mut root, name, 0);
        }
        let mut chunks = vec![root, getitem];
        if old {
            baseline::emit_strided_assign(&mut chunks, 0, 0, baseline::Options::new(false));
        } else {
            slices::emit_strided_assign(&mut chunks, 0, 0, slices::Options::new(false));
        }
        globals::emit_read(&mut chunks[0], "input", 0);
        let chunks = finish(chunks);
        validate(&chunks);
        let input = array((0..8).map(Value::I32).collect());
        // Existing null-valued own keys activate ARRAY_GET's real __getitem__
        // fallback. Dense array imports bypass host overrides in the VM.
        let mut source = Object::new();
        for index in 0..4 {
            source.properties.insert(index.to_string(), Value::Null);
        }
        let source = Value::Object(vybe_runtime::heap::alloc(source));
        let mut vm = runtime(&[
            ("input", input.clone()),
            ("source", source),
            ("lower", Value::I32(0)),
            ("upper", Value::I32(8)),
            ("step", Value::I32(2)),
        ]);
        let trace = Arc::new(Mutex::new(Vec::new()));
        let observed = trace.clone();
        let observed_input = input.clone();
        vm.register_host_fn(
            "ecma:string",
            "charAt",
            Box::new(move |_, args| {
                assert_eq!(args[0], Value::String("source".into()));
                let index = args[1].as_i32();
                observed
                    .lock()
                    .unwrap()
                    .push((index, vector(&observed_input)));
                Value::String(format!("value{index}").into())
            }),
        );
        let result = vm.run(chunks).map(|value| snapshot(&value)).map_err(error);
        let calls = trace.lock().unwrap().clone();
        (result, vector(&input), calls)
    });
    assert_eq!(outcomes[0], outcomes[1]);
    assert_eq!(
        outcomes[0]
            .2
            .iter()
            .map(|(index, _)| *index)
            .collect::<Vec<_>>(),
        vec![0, 1, 2, 3]
    );
    for (index, state) in &outcomes[0].2 {
        for previous in 0..*index {
            assert_eq!(
                state[previous as usize * 2],
                Value::String(format!("value{previous}").into())
            );
        }
        assert_eq!(state[*index as usize * 2], Value::I32(*index * 2));
    }
}
#[test]
fn stepped_string_keeps_character_callback_order_and_count() {
    let outcomes = [program(true, 1, false), program(false, 1, false)].map(|chunks| {
        let mut vm = runtime(&[
            ("input", Value::String("abcdefgh".into())),
            ("lower", Value::I32(0)),
            ("upper", Value::I32(8)),
            ("step", Value::I32(2)),
        ]);
        let trace = Arc::new(Mutex::new(Vec::new()));
        let observed = trace.clone();
        vm.register_host_fn(
            "ecma:string",
            "charAt",
            Box::new(move |_, args| {
                let index = args[1].as_i32() as usize;
                observed.lock().unwrap().push(index);
                Value::String(
                    args[0]
                        .as_str()
                        .chars()
                        .nth(index)
                        .unwrap()
                        .to_string()
                        .into(),
                )
            }),
        );
        let result = vm.run(chunks).map(|value| snapshot(&value)).map_err(error);
        let calls = trace.lock().unwrap().clone();
        (result, calls)
    });
    assert_eq!(outcomes[0], outcomes[1]);
    assert_eq!(outcomes[0].1, vec![0, 2, 4, 6]);
}
#[test]
fn slice_predicates_reduce_emission_and_validate_complete_modules() {
    for operation in 0..5 {
        let old = program(true, operation, false);
        let new = program(false, operation, false);
        validate(&old);
        validate(&new);
        let bytes = |c: &[Chunk]| c.iter().map(|c| c.code.len()).sum::<usize>();
        assert!(bytes(&new) < bytes(&old), "op={operation}");
        eprintln!(
            "slice-predicates op={operation} bytecode={}/{} wasm={}/{}",
            bytes(&old),
            bytes(&new),
            vybe_platform_wasm::write_wasm(&old).len(),
            vybe_platform_wasm::write_wasm(&new).len()
        );
    }
}
#[test]
#[ignore = "interleaved opt-level-0 debug diagnostic"]
fn benchmark_slice_predicates_and_private_counters() {
    for operation in [1, 2, 3] {
        let programs = [
            program(true, operation, false),
            program(false, operation, false),
        ];
        for count in [256, 4096] {
            let mut times = [vec![], vec![]];
            for sample in 0..9 {
                for which in [sample % 2, 1 - sample % 2] {
                    let input = if operation == 1 {
                        Value::String("a".repeat(count).into())
                    } else {
                        array((0..count as i32).map(Value::I32).collect())
                    };
                    let source = array((0..count as i32).map(Value::I32).collect());
                    let mut vm = runtime(&[
                        ("input", input),
                        ("source", source),
                        ("lower", Value::I32(0)),
                        ("upper", Value::I32(count as i32)),
                        ("step", Value::I32(2)),
                    ]);
                    let chunks = programs[which].clone();
                    let start = std::time::Instant::now();
                    let _ = vm.run(chunks).unwrap();
                    times[which].push(start.elapsed().as_micros());
                }
            }
            for time in &mut times {
                time.sort_unstable();
            }
            eprintln!(
                "slice-predicates op={operation} count={count} before={}us after={}us ratio={:.2}x",
                times[0][4],
                times[1][4],
                times[0][4] as f64 / times[1][4] as f64
            );
        }
    }
}
