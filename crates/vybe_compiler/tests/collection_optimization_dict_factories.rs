use vybe_compiler::primitives::{bundle, callable, globals, platforms, polyfills};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/collection_optimization_dict_factories.rs"]
mod baseline;
fn array(values: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)))
}
fn items(value: &Value) -> Vec<Value> {
    let Value::Object(object) = value else {
        panic!("expected array")
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(values) = &object.kind else {
        panic!("expected array")
    };
    values.clone()
}
fn program(old: bool, setdefault: bool, calls: usize) -> Vec<Chunk> {
    let mut root = Chunk::new("<dictionary-private-predicates>");
    let helper = if old {
        if setdefault {
            baseline::build_setdefault(&mut root)
        } else {
            baseline::build_dict_values_from_entries(&mut root)
        }
    } else {
        let name = if setdefault {
            "__stdlib_setdefault"
        } else {
            "__stdlib_dict_values_from_entries"
        };
        polyfills::build_runtime_helpers_for_exports(&mut root, &[name])
            .chunks
            .remove(0)
    };
    let mut chunks = vec![root, helper];
    for call in 0..calls {
        chunks[0].emit_op_u16(Op::REF_FUNC, 1, 0);
        chunks[0].emit(0u8, 0);
        globals::emit_read(&mut chunks[0], "input", 0);
        if setdefault {
            globals::emit_read(&mut chunks[0], "key", 0);
            globals::emit_read(&mut chunks[0], "default", 0);
        }
        callable::emit_direct_invoke_chunk(&mut chunks[0], if setdefault { 3 } else { 1 }, 0);
        if call + 1 < calls {
            chunks[0].emit_op(Op::DROP, 0);
        }
    }
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
fn vm(input: Value, key: Value, default: Value) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("input", input);
    vm.set_global("key", key);
    vm.set_global("default", default);
    vm
}
#[test]
fn dictionary_value_cursor_keeps_order_values_and_fresh_results() {
    let programs = [program(true, false, 1), program(false, false, 1)];
    for chunks in &programs {
        validate(chunks);
    }
    for values in [
        vec![],
        vec![Value::Null],
        vec![
            Value::I32(0),
            Value::Bool(false),
            Value::String("".into()),
            Value::String("é😀".into()),
            Value::Undefined,
        ],
    ] {
        let entries = array(
            values
                .iter()
                .enumerate()
                .map(|(index, value)| array(vec![Value::I32(index as i32), value.clone()]))
                .collect(),
        );
        for chunks in &programs {
            let result = vm(entries.clone(), Value::Null, Value::Null)
                .run(chunks.clone())
                .unwrap();
            assert_eq!(items(&result), values);
            assert_ne!(result, entries);
        }
    }
}
#[test]
fn dictionary_value_cursor_preserves_non_entry_input_outcomes() {
    let programs = [program(true, false, 1), program(false, false, 1)];
    let mut negative_length = Object::new();
    negative_length
        .properties
        .insert("length".into(), Value::I32(-1));
    // The production caller supplies Object.entries arrays. Keep the captured
    // factory's existing outcomes for unsupported receivers and malformed rows.
    for input in [
        Value::Null,
        Value::Undefined,
        Value::Bool(false),
        Value::I32(7),
        Value::Object(vybe_runtime::heap::alloc(negative_length)),
        array(vec![Value::Null, array(vec![Value::I32(0)])]),
        Value::String("é😀".into()),
    ] {
        let outcomes = programs.each_ref().map(|chunks| {
            vm(input.clone(), Value::Null, Value::Null)
                .run(chunks.clone())
                .map(|result| items(&result))
                .map_err(|error| format!("{error:?}"))
        });
        assert_eq!(outcomes[0], outcomes[1], "input={input:?}");
    }
}
fn dictionary(map: bool, value: Value) -> Value {
    let mut object = Object::new();
    if map {
        object.kind = ObjectKind::Map(Default::default());
        if let ObjectKind::Map(entries) = &mut object.kind {
            entries.insert(Value::String("present".into()), value);
        }
    } else {
        object.properties.insert("present".into(), value);
    }
    Value::Object(vybe_runtime::heap::alloc(object))
}
fn dictionary_get(value: &Value, key: &str) -> Value {
    let Value::Object(object) = value else {
        unreachable!()
    };
    let object = object.lock().unwrap();
    if let ObjectKind::Map(map) = &object.kind {
        map.get(&Value::String(key.into()))
            .cloned()
            .unwrap_or(Value::Undefined)
    } else {
        object
            .properties
            .get(key)
            .cloned()
            .unwrap_or(Value::Undefined)
    }
}
#[test]
fn setdefault_preserves_raw_null_overwrite_false_values_and_default_identity() {
    let programs = [program(true, true, 1), program(false, true, 1)];
    for chunks in &programs {
        validate(chunks);
    }
    for map in [false, true] {
        for existing in [
            Value::Null,
            Value::Undefined,
            Value::Bool(false),
            Value::I32(0),
            Value::String("".into()),
            Value::I32(7),
        ] {
            for key in ["present", "missing"] {
                for chunks in &programs {
                    let input = dictionary(map, existing.clone());
                    let default = array(vec![Value::I32(99)]);
                    let expected =
                        if key == "missing" || matches!(existing, Value::Null | Value::Undefined) {
                            default.clone()
                        } else {
                            existing.clone()
                        };
                    let result = vm(input.clone(), Value::String(key.into()), default)
                        .run(chunks.clone())
                        .unwrap();
                    assert_eq!(
                        result, expected,
                        "map={map} existing={existing:?} key={key}"
                    );
                    assert_eq!(dictionary_get(&input, key), expected);
                }
            }
        }
    }
}
#[test]
fn dict_predicate_repeated_modules_validate_and_report_sizes() {
    for setdefault in [false, true] {
        let old = program(true, setdefault, 20);
        let new = program(false, setdefault, 20);
        validate(&old);
        validate(&new);
        let bytes = |chunks: &[Chunk]| chunks.iter().map(|c| c.code.len()).sum::<usize>();
        assert!(bytes(&new) < bytes(&old));
        eprintln!(
            "dict-factory setdefault={setdefault} bytecode={}/{} wasm={}/{}",
            bytes(&old),
            bytes(&new),
            vybe_platform_wasm::write_wasm(&old).len(),
            vybe_platform_wasm::write_wasm(&new).len()
        );
    }
}
#[test]
#[ignore = "interleaved opt-level-0 debug diagnostic"]
fn benchmark_dictionary_values_private_cursor() {
    let programs = [program(true, false, 1), program(false, false, 1)];
    for count in [256, 8192] {
        let mut timings = [vec![], vec![]];
        for sample in 0..9 {
            for which in [sample % 2, 1 - sample % 2] {
                let entries = array(
                    (0..count)
                        .map(|index| array(vec![Value::I32(index), Value::I32(index)]))
                        .collect(),
                );
                let mut runtime = vm(entries, Value::Null, Value::Null);
                let chunks = programs[which].clone();
                let start = std::time::Instant::now();
                let result = runtime.run(chunks).unwrap();
                timings[which].push(start.elapsed().as_micros());
                assert_eq!(items(&result).len(), count as usize);
            }
        }
        for times in &mut timings {
            times.sort_unstable();
        }
        eprintln!(
            "dict-values count={count} before={}us after={}us ratio={:.2}x",
            timings[0][4],
            timings[1][4],
            timings[0][4] as f64 / timings[1][4] as f64
        );
    }
}
