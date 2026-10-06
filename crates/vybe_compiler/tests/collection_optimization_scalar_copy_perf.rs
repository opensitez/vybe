use std::sync::Arc;
use vybe_compiler::primitives::{bundle, clone, globals, platforms};
use vybe_runtime::value::Object;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/collection_optimization_scalar_copy.rs"]
mod baseline;
#[path = "fixtures/collection_optimization_scalar_copy_duplicate.rs"]
mod duplicate;
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
    let mut root = Chunk::new(format!("scalar-copy-old{old}-force{force}-calls{calls}"));
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
fn scalar_bypass_keeps_nested_values_references_and_freshness() {
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
fn scalar_bypass_scalar_exits_and_shared_body_validate() {
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
        "scalar-copy: bytecode={}/{} wasm={}/{}",
        bytes(&old),
        bytes(&new),
        vybe_platform_wasm::write_wasm(&old).len(),
        vybe_platform_wasm::write_wasm(&new).len()
    );
}
fn function_value() -> Value {
    let mut object = Object::new();
    object.kind = vybe_runtime::value::ObjectKind::Function(vybe_runtime::value::Function {
        name: Some("retained-reference".into()),
        arity: 0,
        chunk_index: 0,
        upvalues: vec![],
    });
    Value::Object(vybe_runtime::heap::alloc(object))
}
fn workload(kind: &str, count: usize) -> Value {
    let mut record = Object::new();
    record
        .properties
        .insert("__value_copy".into(), Value::Bool(true));
    for index in 0..count {
        let value = match kind {
            "scalars" => match index % 5 {
                0 => Value::I32(index as i32),
                1 => Value::F64(index as f64 + 0.5),
                2 => Value::String("é😀".into()),
                3 => Value::Bool(false),
                _ => Value::Null,
            },
            "nested" => object(vec![("number", Value::I32(index as i32))], true),
            "functions" => function_value(),
            "references" => object(vec![("number", Value::I32(index as i32))], false),
            _ => unreachable!(),
        };
        record.properties.insert(format!("field{index}"), value);
    }
    Value::Object(vybe_runtime::heap::alloc(record))
}
#[test]
fn scalar_bypass_preserves_function_identity_and_aggregate_workloads() {
    let programs = [program(true, false, 1), program(false, false, 1)];
    for chunks in &programs {
        validate(chunks);
    }
    for kind in ["scalars", "nested", "functions", "references"] {
        for chunks in &programs {
            let input = workload(kind, 20);
            let result = vm(input.clone()).run(chunks.clone()).unwrap();
            for index in 0..20 {
                let key = format!("field{index}");
                let before = property(&input, &key);
                let after = property(&result, &key);
                if kind == "nested" {
                    assert!(!same_object(&before, &after));
                    assert_eq!(property(&before, "number"), property(&after, "number"));
                } else {
                    assert_eq!(before, after, "kind={kind} key={key}");
                }
            }
        }
    }
}
#[test]
#[ignore = "interleaved opt-level-0 debug diagnostic"]
fn benchmark_scalar_field_recursive_call_bypass() {
    let mut intermediate_root = Chunk::new("<duplicate-predicate-copy>");
    intermediate_root.local_count = 1;
    globals::emit_read(&mut intermediate_root, "input", 0);
    intermediate_root.emit_op_u16(Op::LOCAL_SET, 0, 0);
    let mut intermediate = vec![intermediate_root];
    duplicate::emit_deep_copy(&mut intermediate, 0, 0, false, 0);
    intermediate[0].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut intermediate);
    globals::declare_free_globals(&mut intermediate);
    globals::normalize_global_table(&mut intermediate);
    validate(&intermediate);
    let programs = [
        program(true, false, 1),
        intermediate,
        program(false, false, 1),
    ];
    for (kind, count) in [
        ("scalars", 256),
        ("scalars", 8192),
        ("nested", 256),
        ("nested", 1024),
        ("functions", 256),
        ("functions", 8192),
        ("references", 256),
        ("references", 8192),
    ] {
        let mut timings = [vec![], vec![], vec![]];
        for round in 0..9 {
            for which in [round % 3, (round + 1) % 3, (round + 2) % 3] {
                let input = workload(kind, count);
                let expected = property(&input, &format!("field{}", count - 1));
                let mut runtime = vm(input);
                let chunks = programs[which].clone();
                let started = std::time::Instant::now();
                let result = runtime.run(chunks).unwrap();
                timings[which].push(started.elapsed().as_micros());
                let actual = property(&result, &format!("field{}", count - 1));
                if kind == "nested" {
                    assert_eq!(property(&actual, "number"), property(&expected, "number"));
                } else {
                    assert_eq!(actual, expected);
                }
            }
        }
        for times in &mut timings {
            times.sort_unstable();
        }
        eprintln!(
            "scalar-copy kind={kind} fields={count}: before={}us duplicate={}us after={}us ratio={:.2}x seam={:.2}x",
            timings[0][4],
            timings[1][4],
            timings[2][4],
            timings[0][4] as f64 / timings[2][4] as f64,
            timings[1][4] as f64 / timings[2][4] as f64
        );
    }
}

#[test]
fn public_copy_callers_do_not_pass_private_proof_state_to_named_source_chunks() {
    // Preserve the preexisting by-name helper lookup: a same-named source
    // chunk can be selected by public emission, but receives only public 0/1.
    // The proof state is emitted solely by the generated helper's self-call.
    for implementation in 0..3 {
        for force in [false, true] {
            let mut root = Chunk::new("<copy-source-name-compatibility>");
            root.local_count = 1;
            globals::emit_read(&mut root, "input", 0);
            root.emit_op_u16(Op::LOCAL_SET, 0, 0);
            let mut source = Chunk::new("__vybe_value_copy");
            source.arity = 2;
            source.local_count = 2;
            source.emit_op_u16(Op::LOCAL_GET, 1, 0);
            source.emit_op(Op::RETURN, 0);
            let mut chunks = vec![root, source];
            match implementation {
                0 => baseline::emit_deep_copy(&mut chunks, 0, 0, force, 0),
                1 => duplicate::emit_deep_copy(&mut chunks, 0, 0, force, 0),
                _ => clone::emit_deep_copy(&mut chunks, 0, 0, force, 0),
            }
            chunks[0].emit_op(Op::RETURN, 0);
            bundle::finalize_with_runtime_helpers(&mut chunks);
            globals::declare_free_globals(&mut chunks);
            globals::normalize_global_table(&mut chunks);
            validate(&chunks);
            assert_eq!(
                vm(Value::I32(7)).run(chunks).unwrap(),
                Value::I32(i32::from(force))
            );
        }
    }
}
