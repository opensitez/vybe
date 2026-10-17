use vybe_compiler::primitives::{bundle, dict, globals, platforms, sets, tuples};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Chunk, Op, VM, Value};

#[path = "fixtures/dynamic_dict_predicates.rs"]
mod baseline;

fn finish(mut chunks: Vec<Chunk>) -> Vec<Chunk> {
    chunks[0].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
fn program(old: bool, delete: bool) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<dict-predicate>")];
    globals::emit_read(&mut chunks[0], "input", 0);
    globals::emit_read(&mut chunks[0], "key", 0);
    match (old, delete) {
        (true, false) => baseline::emit_method_has(&mut chunks, 0, 0),
        (false, false) => dict::emit_method_has(&mut chunks, 0, 0),
        (true, true) => baseline::emit_method_delete(&mut chunks, 0, 0),
        (false, true) => dict::emit_method_delete(&mut chunks, 0, 0),
    }
    finish(chunks)
}
fn array(values: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)))
}
fn object(value: Value, tracked: bool) -> Value {
    let mut object = Object::new();
    object.properties.insert("present".into(), value);
    if tracked {
        object.properties.insert(
            "__keys".into(),
            array(vec![Value::String("present".into())]),
        );
    }
    Value::Object(vybe_runtime::heap::alloc(object))
}
fn runtime(input: Value, key: &str) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("input", input);
    vm.set_global("key", Value::String(key.into()));
    vm
}
fn validate(chunks: &[Chunk]) {
    let bytes = vybe_platform_wasm::write_wasm(chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap_or_else(|error| panic!("{}: {error}", chunks[0].name));
}

#[test]
fn missing_dict_keys_are_absent_with_or_without_key_tracking() {
    for tracked in [false, true] {
        for value in [Value::I32(0), Value::Bool(false), Value::Null, Value::Undefined] {
            let chunks = program(false, false);
            validate(&chunks);
            assert_eq!(runtime(object(value.clone(), tracked), "missing").run(chunks).unwrap(),
                Value::Bool(false));
            if tracked {
                assert_eq!(runtime(object(value, true), "present").run(program(false, false)).unwrap(),
                    Value::Bool(true));
            }
        }
    }
}

#[test]
fn dict_predicates_preserve_tracked_nulls_fallbacks_and_delete_results() {
    for tracked in [false, true] {
        for value in [
            Value::Null,
            Value::Undefined,
            Value::Bool(false),
            Value::I32(0),
            Value::String("".into()),
        ] {
            for key in ["present", "missing"] {
                for old in [true, false] {
                    let chunks = program(old, false);
                    if old {
                        assert!(
                            wasmparser::Validator::new_with_features(
                                wasmparser::WasmFeatures::all()
                            )
                            .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
                            .is_err(),
                            "captured void branches carry values"
                        );
                    } else {
                        validate(&chunks);
                    }
                    let expected = key == "present"
                        && (tracked || !matches!(value, Value::Null | Value::Undefined));
                    assert_eq!(
                        runtime(object(value.clone(), tracked), key)
                            .run(chunks)
                            .unwrap(),
                        Value::Bool(expected)
                    );
                }
            }
        }
    }
    for key in ["present", "missing"] {
        for old in [true, false] {
            let input = object(Value::I32(7), true);
            let chunks = program(old, true);
            if old {
                assert!(
                    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                        .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
                        .is_err(),
                    "captured void branch carries a Boolean result"
                );
            } else {
                validate(&chunks);
            }
            assert_eq!(
                runtime(input.clone(), key).run(chunks).unwrap(),
                Value::Bool(key == "present")
            );
            assert_eq!(
                runtime(input, "present")
                    .run(program(false, false))
                    .unwrap(),
                Value::Bool(key != "present")
            );
        }
    }
    for delete in [false, true] {
        let a = program(true, delete);
        let b = program(false, delete);
        let size = |p: &[Chunk]| p.iter().map(|c| c.code.len()).sum::<usize>();
        assert!(size(&b) < size(&a));
        println!(
            "dict delete={delete}: bytecode={}/{} wasm={}/{}",
            size(&a),
            size(&b),
            vybe_platform_wasm::write_wasm(&a).len(),
            vybe_platform_wasm::write_wasm(&b).len()
        );
    }
}

fn tuple() -> Value {
    let Value::Object(object) = array(vec![Value::I32(1), Value::I32(2)]) else {
        unreachable!()
    };
    object
        .lock()
        .unwrap()
        .properties
        .insert(tuples::TUPLE_TAG.into(), Value::Bool(true));
    Value::Object(object)
}
#[test]
fn tuple_predicate_consumers_keep_canonical_keys_and_snapshot_membership() {
    let first = tuple();
    let second = tuple();
    let mut map = Object::new();
    map.kind = ObjectKind::Map(Default::default());
    let map = Value::Object(vybe_runtime::heap::alloc(map));
    let mut vm = runtime(map, "unused");
    vm.set_global("first", first.clone());
    vm.set_global("second", second);
    let mut root = Chunk::new("<tuple-key>");
    root.local_count = 2;
    globals::emit_read(&mut root, "input", 0);
    root.emit_op_u16(Op::LOCAL_SET, 0, 0);
    let mut chunks = vec![root];
    for key in ["first", "second"] {
        globals::emit_read(&mut chunks[0], key, 0);
        chunks[0].emit_op_u16(Op::LOCAL_SET, 1, 0);
        dict::emit_canonical_key_slot(&mut chunks, 0, 0, 1, 0);
    }
    chunks[0].emit_op_u16(Op::LOCAL_GET, 1, 0);
    let chunks = finish(chunks);
    validate(&chunks);
    assert_eq!(vm.run(chunks).unwrap(), first);

    for tagged in [true, false] {
        let first = if tagged {
            tuple()
        } else {
            array(vec![Value::I32(1), Value::I32(2)])
        };
        let second = if tagged {
            tuple()
        } else {
            array(vec![Value::I32(1), Value::I32(2)])
        };
        let mut vm = runtime(Value::Null, "unused");
        vm.set_global("first", first);
        vm.set_global("second", second);
        let mut chunks = vec![Chunk::new("<snapshot-membership>")];
        chunks[0].local_count = 1;
        sets::emit_new(&mut chunks, 0, 0);
        chunks[0].emit_op_u16(Op::LOCAL_SET, 0, 0);
        chunks[0].emit_op_u16(Op::LOCAL_GET, 0, 0);
        globals::emit_read(&mut chunks[0], "first", 0);
        sets::emit_add_snapshot(&mut chunks, 0, 0);
        chunks[0].emit_op(Op::DROP, 0);
        chunks[0].emit_op_u16(Op::LOCAL_GET, 0, 0);
        globals::emit_read(&mut chunks[0], "second", 0);
        sets::emit_has_snapshot(&mut chunks, 0, 0);
        let chunks = finish(chunks);
        validate(&chunks);
        assert_eq!(vm.run(chunks).unwrap(), Value::Bool(tagged));
    }
}
