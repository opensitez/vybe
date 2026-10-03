//! Correctness cases and a repeatable, opt-in emission/execution benchmark.
use vybe_compiler::primitives::{
    bundle, collections, convert, globals, heap, ops, platforms, string_similarity, strings,
};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Chunk, Op, VM, Value};
use vybex as _;

#[path = "fixtures/heap_full_sort.rs"]
mod heap_full_sort;
#[path = "fixtures/insertion_sort.rs"]
mod insertion_sort;
#[path = "fixtures/key_selection_sort.rs"]
mod key_selection_sort;
#[path = "fixtures/queue_full_sort.rs"]
mod queue_full_sort;

fn array(values: &[i32]) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(
        values.iter().copied().map(Value::I32).collect(),
    )))
}

fn program(builder: fn(&mut Chunk) -> Chunk) -> Vec<Chunk> {
    let mut script = Chunk::new("<script>");
    let helper = builder(&mut script);
    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0);
    globals::emit_read(&mut script, "input", 0);
    script.emit_op_u8_u8(Op::CALL_REF, 1, 1, 0);
    script.emit_op(Op::RETURN, 0);
    vec![script, helper]
}

fn vm(input: Value) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("input", input);
    vm
}

fn array_items(value: Value) -> Vec<Value> {
    let Value::Object(object) = value else {
        panic!("expected array")
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(items) = &object.kind else {
        panic!("expected array")
    };
    items.clone()
}

#[test]
fn collection_numeric_control_preserves_dynamic_elements() {
    let input = Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
        Value::Null,
        Value::I32(0),
        Value::Bool(false),
        Value::String("".into()),
        Value::I32(5),
    ])));
    let compact = array_items(
        vm(input.clone())
            .run(program(collections::build_compact))
            .unwrap(),
    );
    assert_eq!(compact.len(), 4);
    assert!(matches!(compact[0], Value::I32(0)));
    assert!(matches!(compact[1], Value::Bool(false)));
    assert!(matches!(&compact[2], Value::String(s) if s.is_empty()));
    assert!(matches!(compact[3], Value::I32(5)));
    let pairs = array_items(
        vm(input)
            .run(program(collections::build_enumerate))
            .unwrap(),
    );
    assert_eq!(pairs.len(), 5);
    for (index, pair) in pairs.into_iter().enumerate() {
        let pair = array_items(pair);
        assert_eq!(pair.len(), 2);
        assert_eq!(pair[0].as_i32(), index as i32);
        if index == 0 {
            assert!(matches!(pair[1], Value::Null));
        }
    }
    for (values, any, all) in [
        (&[][..], 0, 1),
        (&[0, 0][..], 0, 0),
        (&[0, 5][..], 1, 0),
        (&[1, 5][..], 1, 1),
    ] {
        assert_eq!(
            vm(array(values))
                .run(program(collections::build_pyany))
                .unwrap()
                .as_i32(),
            any
        );
        assert_eq!(
            vm(array(values))
                .run(program(collections::build_pyall))
                .unwrap()
                .as_i32(),
            all
        );
        assert_eq!(
            vm(array(values))
                .run(program(collections::build_isempty))
                .unwrap()
                .as_i32(),
            i32::from(values.is_empty())
        );
    }
    let text = Value::String("cab".into());
    let characters = array_items(vm(text).run(program(collections::build_pyiter)).unwrap());
    assert_eq!(
        characters.iter().map(Value::as_str).collect::<Vec<_>>(),
        ["c", "a", "b"]
    );
    let strings = Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
        Value::String("b".into()),
        Value::String("a".into()),
    ])));
    assert_eq!(
        vm(strings.clone())
            .run(program(collections::build_min))
            .unwrap()
            .as_str(),
        "a"
    );
    assert_eq!(
        vm(strings)
            .run(program(collections::build_max))
            .unwrap()
            .as_str(),
        "b"
    );
}

#[test]
fn collection_numeric_control_wasm_binaries_validate() {
    for builder in [
        collections::build_sum as fn(&mut Chunk) -> Chunk,
        collections::build_min,
        collections::build_max,
        collections::build_sorted,
        collections::build_enumerate,
        collections::build_compact,
        collections::build_pyany,
        collections::build_pyall,
        collections::build_pyiter,
        collections::build_sort_in_place,
        collections::build_sort_with_comparator,
        collections::build_isempty,
        collections::build_pynext,
    ] {
        let helper = builder(&mut Chunk::new("imports"));
        let name = helper.name.clone();
        let binary = vybe_platform_wasm::write_wasm(&[helper]);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&binary)
            .unwrap_or_else(|error| panic!("{name}: {error}"));
    }
}

#[test]
fn collection_helpers_preserve_results_and_copy_contract() {
    for values in [
        vec![],
        vec![7],
        vec![3, -2, 3, 0, -9],
        (0..64).rev().collect(),
    ] {
        let input = array(&values);
        let mut sorted = values.clone();
        sorted.sort();
        for (builder, copies) in [
            (collections::build_sorted as fn(&mut Chunk) -> Chunk, true),
            (collections::build_sort_in_place, false),
        ] {
            let original = array(&values);
            let result = vm(original.clone()).run(program(builder)).unwrap();
            let Value::Object(result) = result else {
                panic!("expected array")
            };
            let guard = result.lock().unwrap();
            let ObjectKind::Array(items) = &guard.kind else {
                panic!("expected array")
            };
            assert_eq!(items.iter().map(Value::as_i32).collect::<Vec<_>>(), sorted);
            let Value::Object(original) = original else {
                unreachable!()
            };
            if copies {
                assert!(!std::sync::Arc::ptr_eq(&original, &result));
                let original = original.lock().unwrap();
                let ObjectKind::Array(items) = &original.kind else {
                    unreachable!()
                };
                assert_eq!(items.iter().map(Value::as_i32).collect::<Vec<_>>(), values);
            } else {
                assert!(std::sync::Arc::ptr_eq(&original, &result));
            }
        }
        assert_eq!(
            vm(input.clone())
                .run(program(collections::build_sum))
                .unwrap()
                .as_i32(),
            values.iter().sum::<i32>()
        );
        if !values.is_empty() {
            assert_eq!(
                vm(input.clone())
                    .run(program(collections::build_min))
                    .unwrap()
                    .as_i32(),
                *values.iter().min().unwrap()
            );
            assert_eq!(
                vm(input)
                    .run(program(collections::build_max))
                    .unwrap()
                    .as_i32(),
                *values.iter().max().unwrap()
            );
        }
    }
}

fn comparator_sort_program(
    inline: bool,
    parameter_receiver: bool,
    constant: Option<f64>,
    traps: bool,
) -> Vec<Chunk> {
    let mut script = Chunk::new("<script>");
    if parameter_receiver {
        script.module_receiver_abi = vybe_runtime::chunk::ReceiverAbi::Parameter;
    }
    let helper = collections::build_sort_with_comparator(&mut script);
    if !inline {
        script.emit_op_u16(Op::REF_FUNC, 1, 0);
        script.emit(0, 0);
    }
    globals::emit_read(&mut script, "input", 0);
    script.emit_op_u16(Op::REF_FUNC, if inline { 1 } else { 2 }, 0);
    script.emit(0, 0);
    if inline {
        collections::emit_sort_func(std::slice::from_mut(&mut script), 0, 0);
    } else {
        script.emit_op_u8_u8(Op::CALL_REF, 2, 1, 0);
    }
    script.emit_op(Op::RETURN, 0);
    let comparator = comparator_callback(parameter_receiver, constant, traps);
    let mut chunks = if inline {
        vec![script, comparator]
    } else {
        vec![script, helper, comparator]
    };
    // These fixtures write a shared counter; use the compiler's final global
    // index space before handing the chunks to the binary writer or VM.
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}

fn comparator_callback(parameter_receiver: bool, constant: Option<f64>, traps: bool) -> Chunk {
    let mut comparator = Chunk::new("compare_keys");
    let first = u16::from(parameter_receiver);
    comparator.arity = 2 + first as u8;
    comparator.local_count = 2 + first;
    globals::emit_read(&mut comparator, "comparisons", 0);
    comparator.emit_i32_const(1, 0);
    comparator.emit_op(Op::I32_ADD, 0);
    globals::emit_write(&mut comparator, "comparisons", 0);
    if traps {
        comparator.emit_op(Op::UNREACHABLE, 0);
    } else if let Some(value) = constant {
        comparator.emit_f64_const(value, 0);
    } else {
        let get = comparator.add_import("ecma:array", "get");
        for slot in [first, first + 1] {
            comparator.emit_op_u16(Op::LOCAL_GET, slot, 0);
            comparator.emit_i32_const(0, 0);
            comparator.emit_call(get, 2, 0);
        }
        comparator.emit_op(Op::F64_SUB, 0);
    }
    comparator.emit_op(Op::RETURN, 0);
    comparator
}

#[test]
fn merge_sort_preserves_stability_identity_receiver_abi_and_complexity() {
    for inline in [false, true] {
        for parameter_receiver in [false, true] {
            let pairs: Vec<_> = (0..129).map(|id| ((id * 37) % 7, id)).collect();
            let input = Value::Object(vybe_runtime::heap::alloc(Object::new_array(
                pairs.iter().map(|&(key, id)| array(&[key, id])).collect(),
            )));
            let mut runtime = vm(input.clone());
            runtime.set_global("comparisons", Value::I32(0));
            let chunks = comparator_sort_program(inline, parameter_receiver, None, false);
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
                .unwrap();
            let result = runtime.run(chunks).unwrap();
            let (Value::Object(original), Value::Object(sorted)) = (&input, &result) else {
                unreachable!()
            };
            assert!(std::sync::Arc::ptr_eq(original, sorted));
            let actual: Vec<_> = array_items(result)
                .into_iter()
                .map(|pair| {
                    let pair = array_items(pair);
                    (pair[0].as_i32(), pair[1].as_i32())
                })
                .collect();
            let mut expected = pairs;
            expected.sort_by_key(|&(key, _)| key);
            assert_eq!(actual, expected);
            assert!(runtime.global("comparisons").unwrap().as_i32() <= 129 * 9);
        }
    }
}

#[test]
fn merge_sort_handles_uneven_runs_equal_nan_and_trapping_comparators() {
    for size in [2, 3, 7, 8, 9, 31, 33, 127, 129] {
        let values: Vec<i32> = (0..size).map(|i| (i * 31) % 17 - 8).collect();
        let mut expected = values.clone();
        expected.sort();
        for builder in [
            collections::build_sorted as fn(&mut Chunk) -> Chunk,
            collections::build_sort_in_place,
        ] {
            let result = vm(array(&values)).run(program(builder)).unwrap();
            assert_eq!(
                array_items(result)
                    .iter()
                    .map(Value::as_i32)
                    .collect::<Vec<_>>(),
                expected
            );
        }
    }
    for constant in [0.0, f64::NAN] {
        let input = array(&[5, 2, 9, 1]);
        let mut runtime = vm(input);
        runtime.set_global("comparisons", Value::I32(0));
        let result = runtime
            .run(comparator_sort_program(false, false, Some(constant), false))
            .unwrap();
        assert_eq!(
            array_items(result)
                .iter()
                .map(Value::as_i32)
                .collect::<Vec<_>>(),
            [5, 2, 9, 1]
        );
        assert_eq!(runtime.global("comparisons").unwrap().as_i32(), 3);
    }
    let mut runtime = vm(array(&[2, 1]));
    runtime.set_global("comparisons", Value::I32(0));
    assert!(
        runtime
            .run(comparator_sort_program(false, false, None, true))
            .is_err()
    );
}

fn key_sort_program(parameter_receiver: bool, traps: bool) -> Vec<Chunk> {
    key_sort_program_with_emitter(
        parameter_receiver,
        traps,
        true,
        collections::emit_sort_by_key_in_place,
    )
}

fn key_sort_program_with_emitter(
    parameter_receiver: bool,
    traps: bool,
    check_order: bool,
    emitter: fn(&mut [Chunk], usize, u32),
) -> Vec<Chunk> {
    let mut script = Chunk::new("<key-sort>");
    if parameter_receiver {
        script.module_receiver_abi = vybe_runtime::chunk::ReceiverAbi::Parameter;
    }
    globals::emit_read(&mut script, "input", 0);
    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0);
    emitter(std::slice::from_mut(&mut script), 0, 0);
    script.emit_op(Op::RETURN, 0);
    let key = key_callback(parameter_receiver, traps, check_order);
    let mut chunks = vec![script, key];
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}

fn key_callback(parameter_receiver: bool, traps: bool, check_order: bool) -> Chunk {
    let mut key = Chunk::new("key");
    let first = u16::from(parameter_receiver);
    key.arity = 1 + first as u8;
    key.local_count = 1 + first;
    if check_order {
        // The original id must equal the number of keys evaluated so far. This
        // checks input order as well as exactly-once evaluation, including ties.
        key.emit_op_u16(Op::LOCAL_GET, first, 0);
        key.emit_i32_const(1, 0);
        key.emit_op(Op::ARRAY_GET, 0);
        globals::emit_read(&mut key, "key_calls", 0);
        key.emit_op(Op::I32_NE, 0);
        key.emit_if(0);
        key.emit_op(Op::UNREACHABLE, 0);
        key.emit_end(0);
    }
    if traps {
        globals::emit_read(&mut key, "key_calls", 0);
        key.emit_i32_const(2, 0);
        key.emit_op(Op::I32_EQ, 0);
        key.emit_if(0);
        key.emit_op(Op::UNREACHABLE, 0);
        key.emit_end(0);
    }
    globals::emit_read(&mut key, "key_calls", 0);
    key.emit_i32_const(1, 0);
    key.emit_op(Op::I32_ADD, 0);
    globals::emit_write(&mut key, "key_calls", 0);
    key.emit_op_u16(Op::LOCAL_GET, first, 0);
    key.emit_i32_const(0, 0);
    key.emit_op(Op::ARRAY_GET, 0);
    key.emit_op(Op::RETURN, 0);
    key
}

#[test]
fn key_sort_caches_keys_in_order_is_stable_and_preserves_identity() {
    for parameter_receiver in [false, true] {
        for size in [0, 1, 2, 3, 7, 8, 9, 129] {
            for string_keys in [false, true] {
                let pairs: Vec<_> = (0..size).map(|id| ((id * 37 + 3) % 7, id)).collect();
                let values: Vec<_> = pairs
                    .iter()
                    .map(|&(key, id)| {
                        Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
                            if string_keys {
                                Value::String(format!("key{key}").into())
                            } else {
                                Value::I32(key)
                            },
                            Value::I32(id),
                        ])))
                    })
                    .collect();
                let input =
                    Value::Object(vybe_runtime::heap::alloc(Object::new_array(values.clone())));
                let mut runtime = vm(input.clone());
                runtime.set_global("key_calls", Value::I32(0));
                let chunks = key_sort_program(parameter_receiver, false);
                wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                    .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
                    .unwrap();
                let result = runtime.run(chunks).unwrap();
                let (Value::Object(original), Value::Object(sorted)) = (&input, &result) else {
                    unreachable!()
                };
                assert!(std::sync::Arc::ptr_eq(original, sorted));
                let actual = array_items(result);
                let mut expected = pairs;
                expected.sort_by_key(|&(key, _)| key);
                for (value, &(_, id)) in actual.iter().zip(&expected) {
                    let (Value::Object(value), Value::Object(original)) =
                        (value, &values[id as usize])
                    else {
                        unreachable!()
                    };
                    assert!(std::sync::Arc::ptr_eq(value, original));
                }
                assert_eq!(runtime.global("key_calls").unwrap().as_i32(), size);
            }
        }
    }
}

#[test]
fn throwing_key_does_not_partially_reorder_input() {
    let values: Vec<_> = (0..7).map(|id| array(&[7 - id, id])).collect();
    let input = Value::Object(vybe_runtime::heap::alloc(Object::new_array(values.clone())));
    let mut runtime = vm(input.clone());
    runtime.set_global("key_calls", Value::I32(0));
    assert!(runtime.run(key_sort_program(false, true)).is_err());
    assert_eq!(runtime.global("key_calls").unwrap().as_i32(), 2);
    for (actual, expected) in array_items(input).iter().zip(&values) {
        let (Value::Object(actual), Value::Object(expected)) = (actual, expected) else {
            unreachable!()
        };
        assert!(std::sync::Arc::ptr_eq(actual, expected));
    }
}

fn selection_program(
    n: i32,
    keyed: bool,
    parameter_receiver: bool,
    emitter: fn(&mut [Chunk], usize, u8, u32),
) -> Vec<Chunk> {
    selection_program_with_count(Some(n), keyed, parameter_receiver, emitter)
}

fn selection_program_with_count(
    n: Option<i32>,
    keyed: bool,
    parameter_receiver: bool,
    emitter: fn(&mut [Chunk], usize, u8, u32),
) -> Vec<Chunk> {
    let mut script = Chunk::new("<selection>");
    if parameter_receiver {
        script.module_receiver_abi = vybe_runtime::chunk::ReceiverAbi::Parameter;
    }
    if let Some(n) = n {
        script.emit_i32_const(n, 0);
    } else {
        globals::emit_read(&mut script, "count", 0);
    }
    globals::emit_read(&mut script, "input", 0);
    if keyed {
        script.emit_op_u16(Op::REF_FUNC, 1, 0);
        script.emit(0, 0);
    }
    emitter(
        std::slice::from_mut(&mut script),
        0,
        if keyed { 3 } else { 2 },
        0,
    );
    script.emit_op(Op::RETURN, 0);
    let mut chunks = vec![script];
    if keyed {
        chunks.push(key_callback(parameter_receiver, false, true));
    }
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}

#[test]
fn bounded_selection_preserves_stable_ties_input_identity_and_key_order() {
    for parameter_receiver in [false, true] {
        for size in [0, 1, 2, 7, 17, 129] {
            for n in [-3, 0, 1, 2, 5, 17, 129, 130] {
                for largest in [false, true] {
                    let pairs: Vec<_> = (0..size).map(|id| ((id * 37 + 3) % 7, id)).collect();
                    let values: Vec<_> = pairs.iter().map(|&(key, id)| array(&[key, id])).collect();
                    let input =
                        Value::Object(vybe_runtime::heap::alloc(Object::new_array(values.clone())));
                    let mut runtime = vm(input.clone());
                    runtime.set_global("key_calls", Value::I32(0));
                    let chunks = selection_program(
                        n,
                        true,
                        parameter_receiver,
                        if largest {
                            heap::emit_nlargest
                        } else {
                            heap::emit_nsmallest
                        },
                    );
                    if size == 7 && n == 2 {
                        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                            .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
                            .unwrap();
                    }
                    let result = runtime.run(chunks).unwrap();
                    let (Value::Object(original), Value::Object(selected)) = (&input, &result)
                    else {
                        unreachable!()
                    };
                    assert!(!std::sync::Arc::ptr_eq(original, selected));
                    let actual = array_items(result);
                    let mut expected = pairs;
                    expected.sort_by(|a, b| {
                        if largest {
                            b.0.cmp(&a.0)
                        } else {
                            a.0.cmp(&b.0)
                        }
                    });
                    expected.truncate(n.max(0) as usize);
                    assert_eq!(actual.len(), expected.len());
                    for (value, &(_, id)) in actual.iter().zip(&expected) {
                        let (Value::Object(value), Value::Object(original)) =
                            (value, &values[id as usize])
                        else {
                            unreachable!()
                        };
                        assert!(std::sync::Arc::ptr_eq(value, original));
                    }
                    for (actual, expected) in array_items(input).iter().zip(&values) {
                        let (Value::Object(actual), Value::Object(expected)) = (actual, expected)
                        else {
                            unreachable!()
                        };
                        assert!(std::sync::Arc::ptr_eq(actual, expected));
                    }
                    assert_eq!(
                        runtime.global("key_calls").unwrap().as_i32(),
                        if n > 0 { size } else { 0 }
                    );
                }
            }
        }
    }
}

#[test]
fn bounded_selection_without_keys_keeps_numeric_order_and_does_not_mutate() {
    for largest in [false, true] {
        for n in [-1, 0, 1, 2, 7, 99] {
            let values = [9, -7, 3, 3, 0, -1, 4];
            let input = array(&values);
            let result = vm(input.clone())
                .run(selection_program(
                    n,
                    false,
                    false,
                    if largest {
                        heap::emit_nlargest
                    } else {
                        heap::emit_nsmallest
                    },
                ))
                .unwrap();
            let mut expected = values.to_vec();
            expected.sort();
            if largest {
                expected.reverse();
            }
            expected.truncate(n.max(0) as usize);
            assert_eq!(
                array_items(result)
                    .iter()
                    .map(Value::as_i32)
                    .collect::<Vec<_>>(),
                expected
            );
            assert_eq!(
                array_items(input)
                    .iter()
                    .map(Value::as_i32)
                    .collect::<Vec<_>>(),
                values
            );
        }
    }
}

#[test]
fn scalar_helper_equality_preserves_nan_null_and_unicode() {
    for (value, expected) in [
        ("12.5tail", 12.5_f64),
        ("bad", 0.0),
        ("", 0.0),
        ("-0", -0.0),
        ("Infinity", f64::INFINITY),
    ] {
        let result = vm(Value::String(value.into()))
            .run(program(convert::build_val))
            .unwrap();
        assert_eq!(result.as_f64().to_bits(), expected.to_bits());
    }
    for (input, empty, whitespace) in [
        (Value::Null, 1, 1),
        (Value::String("".into()), 1, 1),
        (Value::String(" \t\u{2003}".into()), 0, 1),
        (Value::String("\u{1f600}".into()), 0, 0),
        (Value::String("0".into()), 0, 0),
    ] {
        assert_eq!(
            vm(input.clone())
                .run(program(strings::build_string_is_null_or_empty))
                .unwrap()
                .as_i32(),
            empty
        );
        assert_eq!(
            vm(input)
                .run(program(strings::build_string_is_null_or_whitespace))
                .unwrap()
                .as_i32(),
            whitespace
        );
    }
}

#[test]
fn two_string_concat_preserves_order_and_unicode_without_spills() {
    for (left, right) in [("", ""), ("", "right"), ("left", ""), ("é", "😀")] {
        let mut chunk = Chunk::new("<script>");
        chunk.emit_string_const(left, 0);
        chunk.emit_string_const(right, 0);
        strings::emit_concat(&mut chunk, 2, 0);
        chunk.emit_op(Op::RETURN, 0);
        assert_eq!(chunk.local_count, 0);
        let result = vm(Value::Null).run(vec![chunk]).unwrap();
        let Value::String(result) = result else {
            panic!("expected string")
        };
        assert_eq!(result.as_ref(), format!("{left}{right}"));
    }
}

#[test]
fn multi_string_concat_preserves_left_to_right_order() {
    for parts in [
        vec!["a", "é", "😀"],
        vec!["", "a", "", "b", "c", "😀", "é", ""],
    ] {
        let mut chunk = Chunk::new("<script>");
        for part in &parts {
            chunk.emit_string_const(part, 0);
        }
        strings::emit_concat(&mut chunk, parts.len(), 0);
        chunk.emit_op(Op::RETURN, 0);
        assert_eq!(chunk.local_count as usize, parts.len() - 2);
        let binary = vybe_platform_wasm::write_wasm(&[chunk.clone()]);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&binary)
            .unwrap();
        assert_eq!(
            vm(Value::Null).run(vec![chunk]).unwrap().as_str(),
            parts.concat()
        );
    }
}

#[test]
fn php_equality_keeps_coercion_and_dynamic_reassignment() {
    static REGISTER: std::sync::Once = std::sync::Once::new();
    REGISTER.call_once(vybe_language_php::register);
    for (body, expected) in [
        ("return '1e2' == '100';", true),
        ("return '0' == false;", true),
        ("return null == false;", true),
        ("return 1 == '1';", true),
        ("return 1 === '1';", false),
        ("return 'abc' == 0;", false),
        ("return [1, 2] == [1, 2];", true),
        ("$x = 1; $x = 'abc'; return $x == 0;", false),
        ("$x = 1; $x = '1'; return $x === 1;", false),
    ] {
        let mut runtime = vm(Value::Null);
        let source = format!("<?php {body}");
        let module = vybe_language_php::parse(&source).unwrap();
        let profile =
            vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
        let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
            .compile(&module)
            .unwrap();
        let result = runtime.run(chunks).unwrap();
        assert_eq!(result.as_i32() != 0, expected, "{body}");
    }
}

fn equality_loop(typed: bool) -> Vec<Chunk> {
    let mut c = Chunk::new("<equality-loop>");
    c.local_count = 2;
    c.emit_i32_const(4096, 0);
    c.emit_op_u16(Op::LOCAL_SET, 0, 0);
    c.emit_i32_const(0, 0);
    c.emit_op_u16(Op::LOCAL_SET, 1, 0);
    let (loop_patch, _) = c.emit_loop_s(0);
    // These values were emitted as numeric literals: no dynamic type inference.
    c.emit_f64_const(42.0, 0);
    c.emit_f64_const(42.0, 0);
    if typed {
        c.emit_op(Op::F64_EQ, 0);
    } else {
        let mut imports = Chunk::new("imports");
        ops::emit_dyn_eq_into(&mut imports, &mut c, 0);
    }
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_i32_const(1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_TEE, 0, 0);
    c.emit_br_if(0, 0);
    c.emit_end(0);
    c.patch_loop(loop_patch);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op(Op::RETURN, 0);
    vec![c]
}

#[test]
fn typed_equality_loop_matches_dynamic_result() {
    for typed in [false, true] {
        assert_eq!(
            vm(Value::Null).run(equality_loop(typed)).unwrap().as_i32(),
            4096
        );
    }
}

fn similarity_program(emitter: fn(&mut [Chunk], usize, u8, u32), args: &[&str]) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<similarity>")];
    for value in args {
        chunks[0].emit_string_const(value, 0);
    }
    emitter(&mut chunks, 0, args.len() as u8, 0);
    chunks[0].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    chunks
}

#[test]
fn string_similarity_known_types_keep_distance_and_phonetic_results() {
    for (a, b, expected) in [
        ("", "", 0),
        ("kitten", "sitting", 3),
        ("", "abcd", 4),
        ("café", "cafe", 1),
        ("same", "same", 0),
        ("abcd", "dcba", 4),
    ] {
        let chunks = similarity_program(string_similarity::emit_levenshtein, &[a, b]);
        assert_eq!(vm(Value::Null).run(chunks).unwrap().as_i32(), expected);
    }
    for (input, expected) in [
        ("", ""),
        ("Robert", "R163"),
        ("Rupert", "R163"),
        ("A", "A000"),
        ("Euler", "E460"),
    ] {
        let chunks = similarity_program(string_similarity::emit_soundex, &[input]);
        assert_eq!(vm(Value::Null).run(chunks).unwrap().to_string(), expected);
    }
    // Independently enumerate the Soundex groups, including both bounds misses.
    for letter in b'@'..=b'[' {
        let digit = match letter {
            b'B' | b'F' | b'P' | b'V' => '1',
            b'C' | b'G' | b'J' | b'K' | b'Q' | b'S' | b'X' | b'Z' => '2',
            b'D' | b'T' => '3',
            b'L' => '4',
            b'M' | b'N' => '5',
            b'R' => '6',
            _ => '0',
        };
        let input = format!("A{}", letter as char);
        let chunks = similarity_program(string_similarity::emit_soundex, &[&input]);
        assert_eq!(
            vm(Value::Null).run(chunks).unwrap().to_string(),
            format!("A{digit}00")
        );
    }
}

#[test]
#[ignore = "manual benchmark: run with --ignored --nocapture --test-threads=1"]
fn similarity_emission_benchmark() {
    for (name, emitter, args) in [
        (
            "levenshtein",
            string_similarity::emit_levenshtein as fn(&mut [Chunk], usize, u8, u32),
            vec!["abcdefghijklmnopqrstuvwx", "xwvutsrqponmlkjihgfedcba"],
        ),
        ("soundex", string_similarity::emit_soundex, vec!["Pfister"]),
    ] {
        let chunks = similarity_program(emitter, &args);
        let mut samples = Vec::new();
        for _ in 0..15 {
            let mut vm = vm(Value::Null);
            let code = chunks.clone();
            let start = std::time::Instant::now();
            std::hint::black_box(vm.run(code).unwrap());
            samples.push(start.elapsed());
        }
        samples.sort();
        println!(
            "{name}: bytes={} total_bytes={} locals={} imports={} median_us={}",
            chunks[0].code.len(),
            chunks.iter().map(|c| c.code.len()).sum::<usize>(),
            chunks[0].local_count,
            chunks[0].imports.len(),
            samples[7].as_micros()
        );
    }
}

fn optimized_wasm_fixtures() -> [(&'static str, Vec<Chunk>); 10] {
    [
        (
            "levenshtein",
            similarity_program(string_similarity::emit_levenshtein, &["kitten", "sitting"]),
        ),
        (
            "soundex",
            similarity_program(string_similarity::emit_soundex, &["Robert"]),
        ),
        ("val", program(convert::build_val)),
        ("empty", program(strings::build_string_is_null_or_empty)),
        (
            "whitespace",
            program(strings::build_string_is_null_or_whitespace),
        ),
        ("sorted", program(collections::build_sorted)),
        ("typed-equality", equality_loop(true)),
        ("key-sort", key_sort_program(false, false)),
        (
            "nsmallest",
            selection_program(3, true, false, heap::emit_nsmallest),
        ),
        (
            "nlargest",
            selection_program(3, true, true, heap::emit_nlargest),
        ),
    ]
}

#[test]
fn optimized_wasm_binaries_validate() {
    for (name, chunks) in optimized_wasm_fixtures() {
        let binary = vybe_platform_wasm::write_wasm(&chunks);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&binary)
            .unwrap_or_else(|error| panic!("{name}: {error}"));
    }
}

#[test]
#[ignore = "manual export for wasm-tools validation"]
fn export_optimized_wasm_fixtures() {
    let dir = std::env::temp_dir().join("vybe-primitives-wasm");
    std::fs::create_dir_all(&dir).unwrap();
    for (name, chunks) in optimized_wasm_fixtures() {
        let path = dir.join(format!("{name}.wasm"));
        std::fs::write(&path, vybe_platform_wasm::write_wasm(&chunks)).unwrap();
        println!("{}", path.display());
    }
}

#[test]
#[ignore = "manual benchmark: run with --ignored --nocapture --test-threads=1"]
fn equality_emission_benchmark() {
    for typed in [false, true] {
        let chunks = equality_loop(typed);
        let mut samples = Vec::new();
        for _ in 0..15 {
            let mut vm = vm(Value::Null);
            let code = chunks.clone();
            let start = std::time::Instant::now();
            assert_eq!(std::hint::black_box(vm.run(code).unwrap()).as_i32(), 4096);
            samples.push(start.elapsed());
        }
        samples.sort();
        println!(
            "equality typed={typed}: bytes={} locals={} imports={} median_us={}",
            chunks[0].code.len(),
            chunks[0].local_count,
            chunks[0].imports.len(),
            samples[7].as_micros()
        );
    }
}

#[test]
#[ignore = "manual benchmark: run with --ignored --nocapture --test-threads=1"]
fn collection_emission_benchmark() {
    for (name, builder, count) in [
        (
            "sum",
            collections::build_sum as fn(&mut Chunk) -> Chunk,
            4096,
        ),
        ("min", collections::build_min, 4096),
        ("sorted", collections::build_sorted, 128),
    ] {
        let chunks = program(builder);
        let values: Vec<_> = (0..count).rev().collect();
        let mut samples = Vec::new();
        for _ in 0..9 {
            let mut vm = vm(array(&values));
            let code = chunks.clone();
            let start = std::time::Instant::now();
            std::hint::black_box(vm.run(code).unwrap());
            samples.push(start.elapsed());
        }
        samples.sort();
        println!(
            "{name}: bytes={} locals={} imports={} median_us={}",
            chunks[1].code.len(),
            chunks[1].local_count,
            chunks[1].imports.len(),
            samples[4].as_micros()
        );
    }
}

#[test]
#[ignore = "manual algorithm benchmark: run with --ignored --nocapture --test-threads=1"]
fn sort_algorithm_benchmark() {
    for count in [128, 1024] {
        for distribution in ["reverse", "mixed", "sorted"] {
            let values: Vec<i32> = (0..count)
                .map(|i| match distribution {
                    "reverse" => count - i,
                    "mixed" => (i * 719) % count,
                    _ => i,
                })
                .collect();
            let programs = [
                program(insertion_sort::build_sorted),
                program(collections::build_sorted),
            ];
            let mut samples = [Vec::new(), Vec::new()];
            let mut expected = values.clone();
            expected.sort();
            // Interleave the two algorithms to reduce concurrent-load bias.
            for _ in 0..3 {
                for (index, code) in programs.iter().enumerate() {
                    let mut runtime = vm(array(&values));
                    let code = code.clone();
                    let start = std::time::Instant::now();
                    let result = runtime.run(code).unwrap();
                    samples[index].push(start.elapsed());
                    assert_eq!(
                        array_items(result)
                            .iter()
                            .map(Value::as_i32)
                            .collect::<Vec<_>>(),
                        expected
                    );
                }
            }
            for samples in &mut samples {
                samples.sort();
            }
            println!(
                "n={count} {distribution}: insertion_us={} merge_us={} speedup={:.2} bytes={}/{}",
                samples[0][1].as_micros(),
                samples[1][1].as_micros(),
                samples[0][1].as_secs_f64() / samples[1][1].as_secs_f64(),
                programs[0][1].code.len(),
                programs[1][1].code.len()
            );
        }
    }
}

#[test]
#[ignore = "execution benchmark; run explicitly with --include-ignored --nocapture"]
fn key_sort_algorithm_benchmark() {
    use std::time::Instant;
    for size in [128, 512] {
        let mut samples = [Vec::new(), Vec::new()];
        let programs = [
            key_sort_program_with_emitter(
                false,
                false,
                false,
                key_selection_sort::emit_sort_by_key_in_place,
            ),
            key_sort_program_with_emitter(
                false,
                false,
                false,
                collections::emit_sort_by_key_in_place,
            ),
        ];
        for _ in 0..3 {
            for (which, chunks) in programs.iter().enumerate() {
                let values: Vec<_> = (0..size).map(|id| array(&[size - id, id])).collect();
                let input = Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)));
                let mut runtime = vm(input);
                runtime.set_global("key_calls", Value::I32(0));
                let start = Instant::now();
                let result = runtime.run(chunks.clone()).unwrap();
                samples[which].push(start.elapsed().as_micros());
                assert_eq!(
                    runtime.global("key_calls").unwrap().as_i32(),
                    if which == 0 { size * (size - 1) } else { size }
                );
                let actual: Vec<_> = array_items(result)
                    .into_iter()
                    .map(|pair| array_items(pair)[0].as_i32())
                    .collect();
                assert_eq!(actual, (1..=size).collect::<Vec<_>>());
            }
        }
        for sample in &mut samples {
            sample.sort_unstable();
        }
        println!(
            "key-sort n={size}: selection={}us cached-merge={}us speedup={:.2}x key_calls={} -> {} bytecode={} -> {}",
            samples[0][1],
            samples[1][1],
            samples[0][1] as f64 / samples[1][1] as f64,
            size * (size - 1),
            size,
            programs[0].iter().map(|c| c.code.len()).sum::<usize>(),
            programs[1].iter().map(|c| c.code.len()).sum::<usize>()
        );
    }
}

#[test]
#[ignore = "execution benchmark; run explicitly with --include-ignored --nocapture"]
fn bounded_selection_algorithm_benchmark() {
    use std::time::Instant;
    for (n, largest) in [(1, false), (8, false), (32, false), (8, true)] {
        let size = 4096;
        let programs = [
            selection_program(
                n,
                true,
                false,
                if largest {
                    heap_full_sort::emit_nlargest
                } else {
                    heap_full_sort::emit_nsmallest
                },
            ),
            selection_program(
                n,
                true,
                false,
                if largest {
                    heap::emit_nlargest
                } else {
                    heap::emit_nsmallest
                },
            ),
        ];
        let mut samples = [Vec::new(), Vec::new()];
        for _ in 0..3 {
            for (which, chunks) in programs.iter().enumerate() {
                // A permutation prevents either sorter taking a presorted path.
                let values: Vec<_> = (0..size)
                    .map(|id| array(&[(id * 997) % size, id]))
                    .collect();
                let input = Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)));
                let mut runtime = vm(input);
                runtime.set_global("key_calls", Value::I32(0));
                let start = Instant::now();
                let result = runtime.run(chunks.clone()).unwrap();
                samples[which].push(start.elapsed().as_micros());
                assert_eq!(runtime.global("key_calls").unwrap().as_i32(), size);
                let actual: Vec<_> = array_items(result)
                    .into_iter()
                    .map(|pair| array_items(pair)[0].as_i32())
                    .collect();
                assert_eq!(
                    actual,
                    if largest {
                        (size - n..size).rev().collect::<Vec<_>>()
                    } else {
                        (0..n).collect::<Vec<_>>()
                    }
                );
            }
        }
        for sample in &mut samples {
            sample.sort_unstable();
        }
        println!(
            "bounded-selection size={size} n={n} largest={largest}: full-sort={}us bounded-heap={}us speedup={:.2}x bytecode={} -> {}",
            samples[0][1],
            samples[1][1],
            samples[0][1] as f64 / samples[1][1] as f64,
            programs[0].iter().map(|c| c.code.len()).sum::<usize>(),
            programs[1].iter().map(|c| c.code.len()).sum::<usize>()
        );
    }
}

#[test]
fn null_selection_key_uses_values_without_invoking_a_callback() {
    fn with_null_key(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        heap::emit_nsmallest(chunks, current, argc, line);
    }
    let mut runtime = vm(array(&[4, 2, 9, 1, 3]));
    runtime.set_global("key_calls", Value::I32(0));
    let chunks = selection_program(3, true, false, with_null_key);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
        .unwrap();
    let result = runtime.run(chunks).unwrap();
    assert_eq!(
        array_items(result)
            .iter()
            .map(Value::as_i32)
            .collect::<Vec<_>>(),
        [1, 2, 3]
    );
    assert_eq!(runtime.global("key_calls").unwrap().as_i32(), 0);
}

#[test]
fn python_selection_uses_shared_primitive_and_preserves_source_contract() {
    vybe_language_python::register();
    for (body, expected) in [
        (
            "data = [5, 1, 4, 2, 3]\nresult = heapq.nsmallest(2, data)\nreturn result + data",
            vec![1, 2, 5, 1, 4, 2, 3],
        ),
        ("return heapq.nlargest(3, [1, 5, 2, 4, 3])", vec![5, 4, 3]),
        (
            "data = ['bbb', 'a', 'cc']\nresult = heapq.nsmallest(2, data, key=len)\nreturn [len(result[0]), len(result[1]), len(data[0])]",
            vec![1, 2, 3],
        ),
        ("return heapq.nsmallest(-2, [3, 1, 2])", vec![]),
        ("return heapq.nsmallest(2, [3, 1, 2], key=None)", vec![1, 2]),
        (
            "data = [[1, 0], [2, 1], [2, 2], [1, 3]]\nresult = heapq.nlargest(3, data, key=lambda x: x[0])\nreturn [result[0][1], result[1][1], result[2][1], data[0][1], data[1][1]]",
            vec![1, 2, 0, 0, 1],
        ),
    ] {
        let source = format!("import heapq\n{body}\n");
        let module = vybe_language_python::parse(&source).unwrap();
        let profile =
            vybe_compiler::profile::parse_profile(vybe_language_python::profile_source()).unwrap();
        let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
            .compile(&module)
            .unwrap();
        // Ensure this is actually exercising the selector emission.
        assert!(
            chunks.iter().any(|c| c
                .imports
                .iter()
                .any(|import| import.module == "ecma:value" && import.name == "toNumber")),
            "{body}"
        );
        let result = vm(Value::Null).run(chunks).unwrap();
        assert_eq!(
            array_items(result)
                .iter()
                .map(Value::as_i32)
                .collect::<Vec<_>>(),
            expected,
            "{body}"
        );
    }
}

#[test]
fn dynamic_selection_count_is_normalized_once_and_nan_is_empty() {
    for (count, expected) in [
        (Value::F64(f64::NAN), vec![]),
        (Value::Undefined, vec![]),
        (Value::F64(f64::INFINITY), vec![1, 2, 3, 4, 9]),
        (Value::F64(f64::NEG_INFINITY), vec![]),
        (Value::F64(2.9), vec![1, 2]),
        (Value::F64(-2.9), vec![]),
        (Value::String("2".into()), vec![1, 2]),
        (Value::Bool(true), vec![1]),
    ] {
        let mut runtime = vm(array(&[4, 2, 9, 1, 3]));
        runtime.set_global("count", count.clone());
        let chunks = selection_program_with_count(None, false, false, heap::emit_nsmallest);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
            .unwrap();
        let result = runtime.run(chunks).unwrap();
        assert_eq!(
            array_items(result)
                .iter()
                .map(Value::as_i32)
                .collect::<Vec<_>>(),
            expected,
            "count={count}"
        );
    }
}

#[test]
fn bounded_selection_retains_dynamic_string_and_comparable_object_keys() {
    for largest in [false, true] {
        for objects in [false, true] {
            let pairs: Vec<_> = (0..33).map(|id| ((id * 37) % 7, id)).collect();
            let values: Vec<_> = pairs
                .iter()
                .map(|&(key, id)| {
                    let key = if objects {
                        let mut object = Object::new();
                        object
                            .properties
                            .insert("Ticks".into(), Value::F64(key as f64));
                        Value::Object(vybe_runtime::heap::alloc(object))
                    } else {
                        Value::String(format!("key{key}").into())
                    };
                    Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
                        key,
                        Value::I32(id),
                    ])))
                })
                .collect();
            let input = Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)));
            let mut runtime = vm(input);
            runtime.set_global("key_calls", Value::I32(0));
            let chunks = selection_program(
                5,
                true,
                false,
                if largest {
                    heap::emit_nlargest
                } else {
                    heap::emit_nsmallest
                },
            );
            let result = runtime.run(chunks).unwrap();
            let actual: Vec<_> = array_items(result)
                .into_iter()
                .map(|pair| array_items(pair)[1].as_i32())
                .collect();
            let mut expected = pairs;
            expected.sort_by(|a, b| {
                if largest {
                    b.0.cmp(&a.0)
                } else {
                    a.0.cmp(&b.0)
                }
            });
            assert_eq!(
                actual,
                expected.iter().take(5).map(|p| p.1).collect::<Vec<_>>()
            );
            assert_eq!(runtime.global("key_calls").unwrap().as_i32(), 33);
        }
    }
}

fn queue_program(emitter: fn(&mut [Chunk], usize, u8, u32), argc: u8) -> Vec<Chunk> {
    let mut script = Chunk::new("<queue>");
    globals::emit_read(&mut script, "input", 0);
    if argc == 2 {
        globals::emit_read(&mut script, "added", 0);
    }
    emitter(std::slice::from_mut(&mut script), 0, argc, 0);
    script.emit_op(Op::RETURN, 0);
    let mut chunks = vec![script];
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}

fn queue_comparator_program(
    fixed_helper: bool,
    parameter_receiver: bool,
    constant: Option<f64>,
    traps: bool,
    emitter: fn(&mut [Chunk], usize, u8, u32),
) -> Vec<Chunk> {
    let mut script = Chunk::new("<queue-insert>");
    if parameter_receiver {
        script.module_receiver_abi = vybe_runtime::chunk::ReceiverAbi::Parameter;
    }
    let mut chunks;
    if fixed_helper {
        script.emit_op_u16(Op::REF_FUNC, 1, 0);
        script.emit(0, 0);
        globals::emit_read(&mut script, "input", 0);
        globals::emit_read(&mut script, "added", 0);
        script.emit_op_u8_u8(Op::CALL_REF, 2, 1, 0);
        script.emit_op(Op::RETURN, 0);
        let mut helper = Chunk::new("insert_with_declared_comparator");
        helper.arity = 2;
        // Deliberately no local_count preallocation, matching the SPL adapters.
        heap::emit_push_sorted_with_comparator_func(&mut helper, 0, 1, 2, 0);
        helper.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        helper.emit_op(Op::RETURN, 0);
        chunks = vec![script, helper, comparator_callback(false, constant, traps)];
    } else {
        globals::emit_read(&mut script, "input", 0);
        globals::emit_read(&mut script, "added", 0);
        script.emit_op_u16(Op::REF_FUNC, 1, 0);
        script.emit(0, 0);
        emitter(std::slice::from_mut(&mut script), 0, 3, 0);
        script.emit_op(Op::RETURN, 0);
        chunks = vec![
            script,
            comparator_callback(parameter_receiver, constant, traps),
        ];
    }
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}

#[test]
fn queue_binary_insertion_preserves_ties_identity_abi_and_logarithmic_callbacks() {
    for (fixed, parameter_receiver) in [(false, false), (false, true), (true, false)] {
        for size in [0, 1, 2, 7, 127, 1024] {
            for key in [-1, 0, size / 6, size + 1] {
                let mut pairs: Vec<_> = (0..size).map(|id| (id / 3, id)).collect();
                let input = Value::Object(vybe_runtime::heap::alloc(Object::new_array(
                    pairs.iter().map(|&(key, id)| array(&[key, id])).collect(),
                )));
                let added = array(&[key, -1]);
                let mut runtime = vm(input.clone());
                runtime.set_global("added", added.clone());
                runtime.set_global("comparisons", Value::I32(0));
                let chunks = queue_comparator_program(
                    fixed,
                    parameter_receiver,
                    None,
                    false,
                    heap::emit_push_with_comparator,
                );
                if size == 7 {
                    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                        .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
                        .unwrap();
                }
                runtime.run(chunks).unwrap();
                let sorted = runtime.global("input").unwrap();
                let (Value::Object(original), Value::Object(sorted_object)) = (&input, &sorted)
                else {
                    unreachable!()
                };
                assert!(std::sync::Arc::ptr_eq(original, sorted_object));
                let actual: Vec<_> = array_items(sorted.clone())
                    .into_iter()
                    .map(|pair| {
                        let pair = array_items(pair);
                        (pair[0].as_i32(), pair[1].as_i32())
                    })
                    .collect();
                pairs.push((key, -1));
                pairs.sort_by_key(|p| p.0);
                assert_eq!(actual, pairs);
                let bound = if size == 0 {
                    0
                } else {
                    (size as f64 + 1.0).log2().ceil() as i32 + 1
                };
                assert!(runtime.global("comparisons").unwrap().as_i32() <= bound);
            }
        }
    }
}

#[test]
fn queue_heapification_and_push_use_numeric_order_not_string_order() {
    let input = array(&[10, 2, 100, -1, 20, 3]);
    let mut runtime = vm(input.clone());
    runtime.run(queue_program(heap::emit_heapify, 1)).unwrap();
    assert_eq!(
        array_items(input.clone())
            .iter()
            .map(Value::as_i32)
            .collect::<Vec<_>>(),
        [-1, 2, 3, 10, 20, 100]
    );
    runtime.set_global("added", Value::I32(11));
    runtime.run(queue_program(heap::emit_push, 2)).unwrap();
    assert_eq!(
        array_items(input)
            .iter()
            .map(Value::as_i32)
            .collect::<Vec<_>>(),
        [-1, 2, 3, 10, 11, 20, 100]
    );
}

#[test]
fn queue_comparator_nan_zero_and_errors_preserve_expected_state() {
    for constant in [0.0, f64::NAN] {
        let input = array(&[5, 2, 9]);
        let mut runtime = vm(input.clone());
        runtime.set_global("added", Value::I32(1));
        runtime.set_global("comparisons", Value::I32(0));
        runtime
            .run(queue_comparator_program(
                false,
                false,
                Some(constant),
                false,
                heap::emit_push_with_comparator,
            ))
            .unwrap();
        assert_eq!(
            array_items(input)
                .iter()
                .map(Value::as_i32)
                .collect::<Vec<_>>(),
            [5, 2, 9, 1]
        );
    }
    let input = array(&[1, 2, 3]);
    let mut runtime = vm(input.clone());
    runtime.set_global("added", Value::I32(0));
    runtime.set_global("comparisons", Value::I32(0));
    assert!(
        runtime
            .run(queue_comparator_program(
                false,
                false,
                None,
                true,
                heap::emit_push_with_comparator
            ))
            .is_err()
    );
    assert_eq!(
        array_items(input)
            .iter()
            .map(Value::as_i32)
            .collect::<Vec<_>>(),
        [1, 2, 3]
    );
}

#[test]
fn queue_push_pop_only_replaces_on_strict_increase_including_nan() {
    for (values, added, expected, remaining) in [
        (vec![], Value::I32(2), 2.0, vec![]),
        (vec![1, 2, 3], Value::I32(0), 0.0, vec![1, 2, 3]),
        (vec![1, 2, 3], Value::I32(1), 1.0, vec![1, 2, 3]),
        (vec![1, 2, 3], Value::I32(10), 1.0, vec![2, 3, 10]),
        (vec![1, 2, 3], Value::F64(f64::NAN), f64::NAN, vec![1, 2, 3]),
    ] {
        let input = array(&values);
        let mut runtime = vm(input.clone());
        runtime.set_global("added", added);
        let chunks = queue_program(heap::emit_push_pop, 2);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
            .unwrap();
        let result = runtime.run(chunks).unwrap().as_f64();
        assert!(result == expected || (result.is_nan() && expected.is_nan()));
        assert_eq!(
            array_items(input)
                .iter()
                .map(Value::as_i32)
                .collect::<Vec<_>>(),
            remaining
        );
    }
}

#[test]
#[ignore = "execution benchmark; run explicitly with --include-ignored --nocapture"]
fn queue_binary_insertion_benchmark() {
    use std::time::Instant;
    for size in [128, 4096] {
        let programs = [
            queue_comparator_program(
                false,
                false,
                None,
                false,
                queue_full_sort::emit_push_with_comparator,
            ),
            queue_comparator_program(false, false, None, false, heap::emit_push_with_comparator),
        ];
        let mut samples = [Vec::new(), Vec::new()];
        let mut comparisons = [0, 0];
        for _ in 0..5 {
            for (which, chunks) in programs.iter().enumerate() {
                let values: Vec<_> = (0..size).map(|id| array(&[id, id])).collect();
                let input = Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)));
                let mut runtime = vm(input.clone());
                runtime.set_global("added", array(&[size / 2, -1]));
                runtime.set_global("comparisons", Value::I32(0));
                let start = Instant::now();
                runtime.run(chunks.clone()).unwrap();
                samples[which].push(start.elapsed().as_micros());
                comparisons[which] = runtime.global("comparisons").unwrap().as_i32();
                let actual: Vec<_> = array_items(input)
                    .into_iter()
                    .map(|pair| {
                        let pair = array_items(pair);
                        (pair[0].as_i32(), pair[1].as_i32())
                    })
                    .collect();
                let mut expected: Vec<_> = (0..size).map(|id| (id, id)).collect();
                expected.push((size / 2, -1));
                expected.sort_by_key(|p| p.0);
                assert_eq!(actual, expected);
            }
        }
        for sample in &mut samples {
            sample.sort_unstable();
        }
        println!(
            "queue-insert n={size}: full-sort={}us binary={}us speedup={:.2}x callbacks={}/{} bytecode={}/{}",
            samples[0][2],
            samples[1][2],
            samples[0][2] as f64 / samples[1][2] as f64,
            comparisons[0],
            comparisons[1],
            programs[0].iter().map(|c| c.code.len()).sum::<usize>(),
            programs[1].iter().map(|c| c.code.len()).sum::<usize>()
        );
    }
}
