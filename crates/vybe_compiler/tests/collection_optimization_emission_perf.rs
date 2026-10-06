use vybe_compiler::primitives::{
    bundle, callable, globals, platforms, polyfills, sorted_collection, tuples,
};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Chunk, Op, VM, Value};

#[path = "fixtures/collection_optimization_arrays.rs"]
mod old_arrays;
#[path = "fixtures/collection_optimization_sorted.rs"]
mod old_sorted;
#[path = "fixtures/collection_optimization_tuples.rs"]
mod old_tuples;

fn array(values: &[i32]) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(
        values.iter().copied().map(Value::I32).collect(),
    )))
}
fn finish(mut chunks: Vec<Chunk>) -> Vec<Chunk> {
    chunks[0].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
fn run(chunks: Vec<Chunk>, values: &[(&str, Value)]) -> Value {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    for (name, value) in values {
        vm.set_global(*name, value.clone());
    }
    vm.run(chunks).unwrap()
}
fn validate(chunks: &[Chunk]) {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(chunks))
        .unwrap();
}
fn vector(value: &Value) -> Vec<Value> {
    let Value::Object(object) = value else {
        panic!("expected array: {value:?}")
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(values) = &object.kind else {
        panic!("not an array")
    };
    values.clone()
}
fn array_program(old: bool, operation: usize) -> Vec<Chunk> {
    let mut imports = Chunk::new("<array-helper-check>");
    let helper = if old {
        match operation {
            0 => old_arrays::build_array_remove_value(&mut imports),
            1 => old_arrays::build_array_insert_range(&mut imports),
            _ => old_arrays::build_array_set_range(&mut imports),
        }
    } else {
        let export = match operation {
            0 => "__stdlib_array_remove_value",
            1 => "__stdlib_array_insert_range",
            _ => "__stdlib_array_set_range",
        };
        polyfills::build_runtime_helpers_for_exports(&mut imports, &[export])
            .chunks
            .remove(0)
    };
    let mut chunks = vec![imports, helper];
    chunks[0].emit_op_u16(Op::REF_FUNC, 1, 0);
    chunks[0].emit(0u8, 0);
    globals::emit_read(&mut chunks[0], "input", 0);
    globals::emit_read(&mut chunks[0], "index", 0);
    if operation != 0 {
        globals::emit_read(&mut chunks[0], "source", 0);
    }
    callable::emit_direct_invoke_chunk(&mut chunks[0], if operation == 0 { 2 } else { 3 }, 0);
    if operation != 0 {
        chunks[0].emit_op(Op::DROP, 0);
        globals::emit_read(&mut chunks[0], "input", 0);
    }
    finish(chunks)
}
#[test]
fn array_helpers_keep_aliasing_indices_mutations_and_remove_results() {
    for operation in 0..3 {
        let old = array_program(true, operation);
        let new = array_program(false, operation);
        validate(&old);
        validate(&new);
        assert!(
            new.iter().map(|c| c.code.len()).sum::<usize>()
                < old.iter().map(|c| c.code.len()).sum::<usize>()
        );
        for index in [
            Value::I32(0),
            Value::I32(-1),
            Value::F64(1.5),
            Value::String("1".into()),
        ] {
            for source in [vec![], vec![8], vec![8, 9, 10]] {
                for alias in [false, true] {
                    let inputs = || {
                        let input = array(&[1, 2, 3]);
                        let src = if alias { input.clone() } else { array(&source) };
                        vec![("input", input), ("index", index.clone()), ("source", src)]
                    };
                    let before = run(old.clone(), &inputs());
                    let after = run(new.clone(), &inputs());
                    if operation == 0 {
                        assert_eq!(before, after);
                    } else {
                        assert_eq!(vector(&before), vector(&after));
                    }
                }
            }
        }
    }
}
fn repr_program(old: bool) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<tuple-repr-check>")];
    globals::emit_read(&mut chunks[0], "input", 0);
    if old {
        old_tuples::emit_list_string_to_tuple(&mut chunks[0], 0);
    } else {
        tuples::emit_list_string_to_tuple(&mut chunks[0], 0);
    }
    finish(chunks)
}
#[test]
fn tuple_repr_preserves_unicode_empty_and_separator_contract() {
    let old = repr_program(true);
    let new = repr_program(false);
    validate(&old);
    validate(&new);
    assert!(new[0].code.len() < old[0].code.len());
    for (input, expected) in [
        ("[]", "()"),
        ("[1]", "(1,)"),
        ("[1, 2]", "(1, 2)"),
        // Preserve the current VM substring/UTF-16 mismatch in the frozen
        // baseline; correcting supplementary-character slicing is separate.
        ("[é😀]", "(é😀],)"),
        ("[a,b]", "(a,b,)"),
        ("[a, b, c]", "(a, b, c)"),
    ] {
        for chunks in [&old, &new] {
            assert_eq!(
                run(chunks.clone(), &[("input", Value::String(input.into()))]),
                Value::String(expected.into())
            );
        }
    }
}
fn range_program(old: bool, mode: u8, inclusive: bool) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<sorted-range-check>")];
    globals::emit_read(&mut chunks[0], "input", 0);
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
    finish(chunks)
}
#[test]
fn sorted_range_filters_and_order_are_preserved_and_exports_validate() {
    for mode in 0..3 {
        for inclusive in [false, true] {
            let old = range_program(true, mode, inclusive);
            let new = range_program(false, mode, inclusive);
            validate(&new);
            for values in [vec![], vec![4, 2, 1, 3, 2], vec![-2, -1, 0, 1]] {
                let inputs = || {
                    [
                        ("input", array(&values)),
                        ("lower", Value::I32(1)),
                        ("upper", Value::I32(3)),
                    ]
                };
                assert_eq!(
                    vector(&run(old.clone(), &inputs())),
                    vector(&run(new.clone(), &inputs()))
                );
            }
        }
    }
}
#[test]
#[ignore = "interleaved opt-level-0 debug diagnostic"]
fn benchmark_array_set_range() {
    use std::time::Instant;
    let old = array_program(true, 2);
    let new = array_program(false, 2);
    let values: Vec<i32> = (0..8192).collect();
    let mut timings = [Vec::new(), Vec::new()];
    for sample in 0..9 {
        for which in if sample % 2 == 0 { [0, 1] } else { [1, 0] } {
            let inputs = [
                ("input", array(&values)),
                ("source", array(&values)),
                ("index", Value::I32(0)),
            ];
            let chunks = if which == 0 { old.clone() } else { new.clone() };
            let started = Instant::now();
            run(chunks, &inputs);
            timings[which].push(started.elapsed().as_micros());
        }
    }
    for times in &mut timings {
        times.sort_unstable();
    }
    eprintln!(
        "array-set-range units=8192 before={}us after={}us ratio={:.2}x",
        timings[0][4],
        timings[1][4],
        timings[0][4] as f64 / timings[1][4] as f64
    );
}

#[test]
fn tuple_slice_tag_guards_preserve_result_identity() {
    for old in [true, false] {
        let mut chunks = vec![Chunk::new("<tuple-propagation-check>")];
        globals::emit_read(&mut chunks[0], "result", 0);
        globals::emit_read(&mut chunks[0], "source", 0);
        if old {
            old_tuples::emit_propagate_tag(&mut chunks, 0, 0);
        } else {
            tuples::emit_propagate_tag(&mut chunks, 0, 0);
        }
        let chunks = finish(chunks);
        if old {
            assert!(
                wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                    .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
                    .is_err(),
                "captured helper consumes result below a nested block entry"
            );
        } else {
            validate(&chunks);
        }
        for tagged in [false, true] {
            for source_array in [false, true] {
                for result_array in [false, true] {
                    let source = if source_array {
                        array(&[1, 2])
                    } else {
                        Value::String("ab".into())
                    };
                    if tagged {
                        if let Value::Object(object) = &source {
                            object
                                .lock()
                                .unwrap()
                                .properties
                                .insert(tuples::TUPLE_TAG.into(), Value::Bool(true));
                        }
                    }
                    let result = if result_array {
                        array(&[2])
                    } else {
                        Value::String("b".into())
                    };
                    let output = run(
                        chunks.clone(),
                        &[("result", result.clone()), ("source", source)],
                    );
                    assert_eq!(output, result);
                    if let Value::Object(object) = &output {
                        assert_eq!(
                            object
                                .lock()
                                .unwrap()
                                .properties
                                .contains_key(tuples::TUPLE_TAG),
                            tagged && source_array && result_array
                        );
                    }
                }
            }
        }
    }
}

fn map_value(descending: bool, empty: bool) -> Value {
    let mut object = Object::new();
    object.kind = ObjectKind::Map(Default::default());
    if !empty {
        if let ObjectKind::Map(map) = &mut object.kind {
            for key in [3, 1, 2] {
                map.insert(Value::I32(key), Value::I32(key * 10));
            }
        }
    }
    object.properties.insert(
        sorted_collection::DESCENDING_MAP_KEY.into(),
        Value::Bool(descending),
    );
    Value::Object(vybe_runtime::heap::alloc(object))
}
fn map_program(old: bool, entries: bool) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<sorted-map-loop-check>")];
    globals::emit_read(&mut chunks[0], "input", 0);
    match (old, entries) {
        (true, false) => old_sorted::emit_sorted_map_values(&mut chunks, 0, 0),
        (false, false) => sorted_collection::emit_sorted_map_values(&mut chunks, 0, 0),
        (true, true) => old_sorted::emit_sorted_map_entries(&mut chunks, 0, 0),
        (false, true) => sorted_collection::emit_sorted_map_entries(&mut chunks, 0, 0),
    }
    finish(chunks)
}
#[test]
fn sorted_map_private_loops_preserve_key_order_values_entries_and_input() {
    for entries in [false, true] {
        let old = map_program(true, entries);
        let new = map_program(false, entries);
        validate(&new);
        for descending in [false, true] {
            for empty in [false, true] {
                let mut expected_keys = if empty { vec![] } else { vec![1, 2, 3] };
                if descending {
                    expected_keys.reverse();
                }
                for chunks in [&old, &new] {
                    let input = map_value(descending, empty);
                    let output = run(chunks.clone(), &[("input", input.clone())]);
                    let actual = vector(&output);
                    if entries {
                        let actual_pairs: Vec<Vec<Value>> = actual.iter().map(vector).collect();
                        let expected: Vec<Vec<Value>> = expected_keys
                            .iter()
                            .map(|k| vec![Value::I32(*k), Value::I32(k * 10)])
                            .collect();
                        assert_eq!(actual_pairs, expected);
                    } else {
                        assert_eq!(
                            actual,
                            expected_keys
                                .iter()
                                .map(|k| Value::I32(k * 10))
                                .collect::<Vec<_>>()
                        );
                    }
                    let Value::Object(object) = input else {
                        unreachable!()
                    };
                    let object = object.lock().unwrap();
                    let ObjectKind::Map(map) = &object.kind else {
                        unreachable!()
                    };
                    assert_eq!(
                        map.keys().cloned().collect::<Vec<_>>(),
                        if empty {
                            vec![]
                        } else {
                            vec![Value::I32(3), Value::I32(1), Value::I32(2)]
                        }
                    );
                }
            }
        }
    }
}
