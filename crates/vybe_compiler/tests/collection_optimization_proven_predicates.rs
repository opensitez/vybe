use std::sync::{Arc, Mutex};
use vybe_compiler::primitives::{
    bundle, callable, globals, platforms, polyfills, sorted_collection,
};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/collection_optimization_reverse_range.rs"]
mod old_reverse;
#[path = "fixtures/collection_optimization_sorted_bool.rs"]
mod old_sorted;
fn array(values: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)))
}
fn vector(value: &Value) -> Vec<Value> {
    let Value::Object(o) = value else {
        panic!("expected array {value:?}")
    };
    let o = o.lock().unwrap();
    let ObjectKind::Array(v) = &o.kind else {
        panic!("expected array")
    };
    v.clone()
}
fn finish(mut chunks: Vec<Chunk>) -> Vec<Chunk> {
    chunks[0].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
fn validate(chunks: &[Chunk]) {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(chunks))
        .unwrap();
}
fn runtime(values: &[(&str, Value)]) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    for (k, v) in values {
        vm.set_global(*k, v.clone());
    }
    vm
}
fn same_error_except_instruction_positions(mut error: vybe_runtime::error::VMError) -> String {
    // Instruction addresses move when emitted bodies shrink. Keep message,
    // frame names/order, source lines, and every other diagnostic field exact.
    for frame in &mut error.call_stack {
        frame.offset = 0;
    }
    format!("{error:?}")
}
fn reverse(old: bool) -> Vec<Chunk> {
    reverse_with_freeze(old, false)
}
fn reverse_with_freeze(old: bool, freeze: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("<reverse-range>");
    let helper = if old {
        old_reverse::build_array_reverse_range(&mut root)
    } else {
        polyfills::build_runtime_helpers_for_exports(&mut root, &["__stdlib_array_reverse_range"])
            .chunks
            .remove(0)
    };
    let mut chunks = vec![root, helper];
    if freeze {
        globals::emit_read(&mut chunks[0], "input", 0);
        let freeze = chunks[0].add_import("ecma:object", "freeze");
        chunks[0].emit_call(freeze, 1, 0);
        chunks[0].emit_op(Op::DROP, 0);
    }
    chunks[0].emit_op_u16(Op::REF_FUNC, 1, 0);
    chunks[0].emit(0, 0);
    for name in ["input", "index", "count"] {
        globals::emit_read(&mut chunks[0], name, 0);
    }
    callable::emit_direct_invoke_chunk(&mut chunks[0], 3, 0);
    finish(chunks)
}
fn sorted(old: bool, add: bool, mode: u8, inclusive: bool, comparator: bool) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<sorted-predicate>")];
    if comparator {
        let mut cmp = Chunk::new("<observed-tie-comparator>");
        cmp.arity = 2;
        cmp.local_count = 2;
        cmp.emit_op_u16(Op::LOCAL_GET, 0, 0);
        cmp.emit_op_u16(Op::LOCAL_GET, 1, 0);
        let call = cmp.add_import("ecma:string", "charAt");
        cmp.emit_call(call, 2, 0);
        cmp.emit_op(Op::RETURN, 0);
        chunks.push(cmp);
        globals::emit_read(&mut chunks[0], "input", 0);
        chunks[0].emit_string_const(sorted_collection::COMPARATOR_KEY, 0);
        chunks[0].emit_op_u16(Op::REF_FUNC, 1, 0);
        chunks[0].emit(0, 0);
        let set = chunks[0].add_import("ecma:object", "set");
        chunks[0].emit_call(set, 3, 0);
        chunks[0].emit_op(Op::DROP, 0);
    }
    globals::emit_read(&mut chunks[0], "input", 0);
    if add {
        globals::emit_read(&mut chunks[0], "value", 0);
        if old {
            old_sorted::emit_sorted_add(&mut chunks, 0, 0);
        } else {
            sorted_collection::emit_sorted_add(&mut chunks, 0, 0);
        }
    } else {
        if mode == 0 || mode == 2 {
            globals::emit_read(&mut chunks[0], "lower", 0);
        }
        if mode == 0 || mode == 1 {
            globals::emit_read(&mut chunks[0], "upper", 0);
        }
        if old {
            old_sorted::emit_sorted_set_range_view(&mut chunks, 0, mode, inclusive, 0);
        } else {
            sorted_collection::emit_sorted_set_range_view(&mut chunks, 0, mode, inclusive, 0);
        }
    }
    finish(chunks)
}
#[test]
fn negative_one_constant_matches_frozen_negation_bits_and_exports() {
    for old in [true, false] {
        let mut imports = Chunk::new("<constant-minus-one>");
        if old {
            imports.emit_i32_const(1, 0);
            let mut dummy = Chunk::new("imports");
            vybe_compiler::primitives::ops::emit_dyn_neg_into(&mut dummy, &mut imports, 0);
        } else {
            imports.emit_f64_const(-1.0, 0);
        }
        let chunks = finish(vec![imports]);
        validate(&chunks);
        let Value::F64(result) = runtime(&[]).run(chunks).unwrap() else {
            panic!("negation must remain F64")
        };
        assert_eq!(result.to_bits(), (-1.0f64).to_bits());
    }
}
#[test]
fn reverse_keeps_dynamic_bounds_mutations_identity_and_errors() {
    let programs = [reverse(true), reverse(false)];
    for chunks in &programs {
        validate(chunks);
    }
    for index in [
        Value::I32(0),
        Value::I32(-1),
        Value::F64(1.5),
        Value::Null,
        Value::Bool(false),
    ] {
        for count in [
            Value::I32(0),
            Value::I32(3),
            Value::F64(2.5),
            Value::String("2".into()),
            Value::Undefined,
        ] {
            let outcomes = programs.each_ref().map(|chunks| {
                let input = array((0..6).map(Value::I32).collect());
                let result = runtime(&[
                    ("input", input.clone()),
                    ("index", index.clone()),
                    ("count", count.clone()),
                ])
                .run(chunks.clone())
                .map_err(same_error_except_instruction_positions);
                (result, vector(&input))
            });
            assert_eq!(outcomes[0], outcomes[1], "index={index:?}, count={count:?}");
        }
    }
    // ecma:array.set's numeric element path currently ignores freeze flags.
    // Preserve that captured behavior, including with the actual freeze API;
    // enforcing frozen numeric writes is a separate platform/runtime change.
    for (marker, freeze) in [("__frozen", false), ("__vybe_frozen", false), ("", true)] {
        let programs = [
            reverse_with_freeze(true, freeze),
            reverse_with_freeze(false, freeze),
        ];
        let outcomes = programs.each_ref().map(|chunks| {
            validate(chunks);
            let input = array((0..6).map(Value::I32).collect());
            if !marker.is_empty() {
                let Value::Object(o) = &input else {
                    unreachable!()
                };
                o.lock()
                    .unwrap()
                    .properties
                    .insert(marker.into(), Value::I32(1));
            }
            let result = runtime(&[
                ("input", input.clone()),
                ("index", Value::I32(0)),
                ("count", Value::I32(5)),
            ])
            .run(chunks.clone())
            .map_err(same_error_except_instruction_positions);
            (result, vector(&input))
        });
        assert_eq!(outcomes[0], outcomes[1], "marker={marker}, freeze={freeze}");
    }
}
#[test]
fn reverse_retains_get_set_order_with_string_indices() {
    let outcomes = [reverse(true), reverse(false)].map(|chunks| {
        let input = array((0..6).map(Value::I32).collect());
        let mut vm = runtime(&[
            ("input", input.clone()),
            ("index", Value::String("1".into())),
            ("count", Value::I32(0)),
        ]);
        let trace = Arc::new(Mutex::new(Vec::new()));
        for setter in [false, true] {
            let observed = trace.clone();
            vm.register_host_fn(
                "ecma:array",
                if setter { "set" } else { "get" },
                Box::new(move |_, args| {
                    observed
                        .lock()
                        .unwrap()
                        .push((setter, format!("{:?}", args[1])));
                    let Value::Object(o) = &args[0] else {
                        unreachable!()
                    };
                    let mut o = o.lock().unwrap();
                    let ObjectKind::Array(values) = &mut o.kind else {
                        unreachable!()
                    };
                    let index = args[1].as_f64() as usize;
                    if setter {
                        if index >= values.len() {
                            values.resize(index + 1, Value::Undefined);
                        }
                        values[index] = args[2].clone();
                        Value::Null
                    } else {
                        values.get(index).cloned().unwrap_or(Value::Undefined)
                    }
                }),
            );
        }
        let result = vm
            .run(chunks)
            .map_err(same_error_except_instruction_positions);
        let calls = trace.lock().unwrap().clone();
        (result, vector(&input), calls)
    });
    assert_eq!(outcomes[0], outcomes[1]);
    assert!(!outcomes[0].2.is_empty());
}
#[test]
fn sorted_bool_consumers_preserve_nan_membership_ranges_and_comparator_calls() {
    for add in [false, true] {
        for mode in 0..3 {
            for inclusive in [false, true] {
                for comparator in [false, true] {
                    let programs = [
                        sorted(true, add, mode, inclusive, comparator),
                        sorted(false, add, mode, inclusive, comparator),
                    ];
                    for chunks in &programs {
                        validate(chunks);
                    }
                    for lower in [Value::I32(1), Value::String("1".into()), Value::Null] {
                        for value in [Value::I32(2), Value::I32(7), Value::F64(f64::NAN)] {
                            let outcomes = programs.each_ref().map(|chunks| {
                                let input =
                                    array(vec![Value::I32(3), Value::I32(1), Value::I32(2)]);
                                let mut vm = runtime(&[
                                    ("input", input.clone()),
                                    ("value", value.clone()),
                                    ("lower", lower.clone()),
                                    ("upper", Value::I32(3)),
                                ]);
                                let trace = Arc::new(Mutex::new(Vec::new()));
                                let observed = trace.clone();
                                vm.register_host_fn(
                                    "ecma:string",
                                    "charAt",
                                    Box::new(move |_, args| {
                                        observed
                                            .lock()
                                            .unwrap()
                                            .push((args[0].clone(), args[1].clone()));
                                        Value::String("0".into())
                                    }),
                                );
                                let output = vm
                                    .run(chunks.clone())
                                    .map(|result| if add { vec![result] } else { vector(&result) })
                                    .map_err(same_error_except_instruction_positions);
                                let calls = trace.lock().unwrap().clone();
                                (output, vector(&input), calls)
                            });
                            // NaN equality is deliberately non-reflexive; compare diagnostic
                            // representations while preserving object identity in other tests.
                            assert_eq!(format!("{:?}",outcomes[0]),format!("{:?}",outcomes[1]),"add={add}, mode={mode}, inclusive={inclusive}, comparator={comparator}");
                        }
                    }
                }
            }
        }
    }
}
#[test]
fn proven_predicates_reduce_opcode_and_native_module_sizes() {
    for (before, after) in [
        (reverse(true), reverse(false)),
        (
            sorted(true, true, 0, true, false),
            sorted(false, true, 0, true, false),
        ),
        (
            sorted(true, false, 0, true, false),
            sorted(false, false, 0, true, false),
        ),
    ] {
        let bytes = |chunks: &[Chunk]| chunks.iter().map(|c| c.code.len()).sum::<usize>();
        assert!(bytes(&after) < bytes(&before));
        validate(&before);
        validate(&after);
        eprintln!(
            "proven-predicates bytecode={}/{} wasm={}/{}",
            bytes(&before),
            bytes(&after),
            vybe_platform_wasm::write_wasm(&before).len(),
            vybe_platform_wasm::write_wasm(&after).len()
        );
    }
}
#[test]
#[ignore = "interleaved opt-level-0 debug diagnostic"]
fn benchmark_reverse_and_range_predicates() {
    for range in [false, true] {
        let programs = if range {
            [
                sorted(true, false, 0, true, false),
                sorted(false, false, 0, true, false),
            ]
        } else {
            [reverse(true), reverse(false)]
        };
        for count in [256, 8192] {
            let mut times = [vec![], vec![]];
            for sample in 0..9 {
                for which in [sample % 2, 1 - sample % 2] {
                    let input = array((0..count).map(Value::I32).collect());
                    let mut vm = runtime(&[
                        ("input", input),
                        ("index", Value::I32(0)),
                        ("count", Value::I32(count)),
                        ("lower", Value::I32(0)),
                        ("upper", Value::I32(count)),
                    ]);
                    let chunks = programs[which].clone();
                    let start = std::time::Instant::now();
                    let result = vm.run(chunks).unwrap();
                    times[which].push(start.elapsed().as_micros());
                    if range {
                        assert_eq!(vector(&result).len(), count as usize);
                    }
                }
            }
            for time in &mut times {
                time.sort_unstable();
            }
            eprintln!(
                "proven-predicates range={range} count={count} before={}us after={}us ratio={:.2}x",
                times[0][4],
                times[1][4],
                times[0][4] as f64 / times[1][4] as f64
            );
        }
    }
}
