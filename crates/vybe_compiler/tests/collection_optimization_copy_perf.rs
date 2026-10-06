use std::sync::Arc;
use vybe_compiler::primitives::{bundle, clone, globals, platforms};
use vybe_runtime::value::Object;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/collection_optimization_copy.rs"]
mod baseline;
fn object(fields: Vec<(&str, Value)>, marked: bool) -> Value {
    let mut object = Object::new();
    for (name, value) in fields {
        object.properties.insert(name.into(), value);
    }
    if marked {
        object
            .properties
            .insert("__value_copy".into(), Value::Bool(true));
    }
    Value::Object(vybe_runtime::heap::alloc(object))
}
fn property(value: &Value, key: &str) -> Value {
    let Value::Object(object) = value else {
        panic!("expected aggregate")
    };
    object
        .lock()
        .unwrap()
        .properties
        .get(key)
        .cloned()
        .unwrap_or(Value::Undefined)
}
fn same_object(a: &Value, b: &Value) -> bool {
    match (a, b) {
        (Value::Object(a), Value::Object(b)) => Arc::ptr_eq(a, b),
        _ => false,
    }
}
fn program(old: bool, force: bool, calls: usize) -> Vec<Chunk> {
    let mut root = Chunk::new(format!("copy-old{old}-force{force}-calls{calls}"));
    root.local_count = 1;
    globals::emit_read(&mut root, "input", 0);
    root.emit_op_u16(Op::LOCAL_SET, 0, 0);
    let mut chunks = vec![root];
    for call in 0..calls {
        if old {
            baseline::emit_deep_copy(&mut chunks, 0, 0, force, 0);
        } else {
            clone::emit_deep_copy(&mut chunks, 0, 0, force, 0);
        }
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
fn vm(input: Value) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("input", input);
    vm
}
fn validate(chunks: &[Chunk]) {
    let bytes = vybe_platform_wasm::write_wasm(chunks);
    if let Err(error) = wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
    {
        let path = format!("/private/tmp/{}.wasm", chunks[0].name);
        std::fs::write(&path, &bytes).unwrap();
        panic!("{}: {error}; diagnostic export {path}", chunks[0].name);
    }
}

#[test]
fn private_key_iteration_keeps_nested_values_references_stamps_and_freshness() {
    for old in [true, false] {
        for force in [false, true] {
            let chunks = program(old, force, 2);
            validate(&chunks);
            for marked in [false, true] {
                let nested = object(vec![("number", Value::I32(7))], true);
                let reference = object(vec![("number", Value::I32(9))], false);
                let input = object(
                    vec![
                        ("nested", nested.clone()),
                        ("reference", reference.clone()),
                        ("zero", Value::I32(0)),
                        ("text", Value::String("é😀".into())),
                        ("null", Value::Null),
                        ("flag", Value::Bool(false)),
                    ],
                    marked,
                );
                let result = vm(input.clone()).run(chunks.clone()).unwrap();
                if !force && !marked {
                    assert!(same_object(&result, &input));
                } else {
                    assert!(!same_object(&result, &input));
                    assert!(!same_object(&property(&result, "nested"), &nested));
                    assert_eq!(
                        property(&property(&result, "nested"), "number"),
                        Value::I32(7)
                    );
                    assert!(same_object(&property(&result, "reference"), &reference));
                    for key in ["zero", "text", "null", "flag"] {
                        assert_eq!(
                            property(&result, key),
                            property(&input, key),
                            "old={old} force={force} marked={marked} key={key}"
                        );
                    }
                    // The captured Object.assign path drops hidden metadata.
                    // Preserve that existing behavior; repairing value-copy
                    // stamps and runtime type metadata is a separate audit.
                    assert_eq!(
                        property(&result, "__value_copy"),
                        Value::Undefined,
                        "old={old} force={force} marked={marked}"
                    );
                    assert_eq!(
                        property(&input, "__value_copy"),
                        if marked {
                            Value::Bool(true)
                        } else {
                            Value::Undefined
                        }
                    );
                    let Value::Object(child) = property(&result, "nested") else {
                        unreachable!()
                    };
                    child
                        .lock()
                        .unwrap()
                        .properties
                        .insert("number".into(), Value::I32(99));
                    assert_eq!(property(&nested, "number"), Value::I32(7));
                }
            }
        }
    }
}
#[test]
fn copy_scalar_exits_empty_aggregates_and_one_shared_body_validate() {
    for old in [true, false] {
        let chunks = program(old, false, 20);
        validate(&chunks);
        assert_eq!(
            chunks
                .iter()
                .filter(|c| c.name == "__vybe_value_copy")
                .count(),
            1
        );
        for input in [
            Value::Null,
            Value::Bool(false),
            Value::I32(0),
            Value::String("text".into()),
        ] {
            assert_eq!(vm(input.clone()).run(chunks.clone()).unwrap(), input);
        }
        let empty = object(vec![], false);
        let forced = program(old, true, 1);
        validate(&forced);
        assert!(!same_object(
            &vm(empty.clone()).run(forced).unwrap(),
            &empty
        ));
    }
    let old = program(true, true, 20);
    let new = program(false, true, 20);
    let bytes = |chunks: &[Chunk]| chunks.iter().map(|c| c.code.len()).sum::<usize>();
    eprintln!(
        "copy: bytecode={}/{} wasm={}/{}",
        bytes(&old),
        bytes(&new),
        vybe_platform_wasm::write_wasm(&old).len(),
        vybe_platform_wasm::write_wasm(&new).len()
    );
}
fn wide_record(count: usize) -> Value {
    let mut object = Object::new();
    object
        .properties
        .insert("__value_copy".into(), Value::Bool(true));
    for index in 0..count {
        object
            .properties
            .insert(format!("field{index}"), Value::I32(index as i32));
    }
    Value::Object(vybe_runtime::heap::alloc(object))
}
#[test]
#[ignore = "interleaved opt-level-0 debug diagnostic"]
fn benchmark_private_copy_key_length_cache() {
    let programs = [program(true, false, 1), program(false, false, 1)];
    for count in [256, 8192] {
        let mut timings = [vec![], vec![]];
        for round in 0..9 {
            for which in [round % 2, 1 - round % 2] {
                let mut runtime = vm(wide_record(count));
                let chunks = programs[which].clone();
                let start = std::time::Instant::now();
                let result = runtime.run(chunks).unwrap();
                timings[which].push(start.elapsed().as_micros());
                assert_eq!(
                    property(&result, &format!("field{}", count - 1)),
                    Value::I32((count - 1) as i32)
                );
            }
        }
        for times in &mut timings {
            times.sort_unstable();
        }
        eprintln!(
            "copy fields={count}: before={}us after={}us ratio={:.2}x",
            timings[0][4],
            timings[1][4],
            timings[0][4] as f64 / timings[1][4] as f64
        );
    }
}
