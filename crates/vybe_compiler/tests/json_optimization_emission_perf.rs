use vybe_compiler::primitives::{bundle, globals, json, platforms};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/json_optimization_baseline.rs"]
mod baseline;

fn program(
    old: bool,
    calls: usize,
    hook: bool,
    sort: bool,
    props: bool,
    own: bool,
    prototype: bool,
) -> Vec<Chunk> {
    program_seeded(old, calls, hook, sort, props, own, prototype, vec![])
}

fn program_seeded(
    old: bool,
    calls: usize,
    hook: bool,
    sort: bool,
    props: bool,
    own: bool,
    prototype: bool,
    extra: Vec<Chunk>,
) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<json-optimization>")];
    chunks.extend(extra);
    chunks[0].module_instance_fields_are_own_properties = own;
    chunks[0].module_members_on_prototype = prototype;
    chunks[0].alloc_scratch(6);
    if hook {
        let mut body = Chunk::new("<json-test-hook>");
        body.arity = 1;
        body.local_count = 1;
        body.emit_op_u16(Op::LOCAL_GET, 0, 0);
        let import = body.add_import("test:json", "hook");
        body.emit_call(import, 1, 0);
        body.emit_op(Op::RETURN, 0);
        chunks.push(body);
    }
    for call in 0..calls {
        globals::emit_read(&mut chunks[0], "input", 0);
        chunks[0].emit_op_u16(Op::LOCAL_SET, 0, 0);
        if hook {
            chunks[0].emit_op_u16(Op::REF_FUNC, 1, 0);
            chunks[0].emit(0, 0);
        } else {
            chunks[0].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        }
        chunks[0].emit_op_u16(Op::LOCAL_SET, 1, 0);
        chunks[0].emit_bool_const(sort, 0);
        chunks[0].emit_op_u16(Op::LOCAL_SET, 2, 0);
        chunks[0].emit_bool_const(props, 0);
        chunks[0].emit_op_u16(Op::LOCAL_SET, 3, 0);
        if old {
            baseline::emit_normalize(&mut chunks, 0, 0, 1, 2, 3, 0);
        } else {
            json::emit_normalize(&mut chunks, 0, 0, 1, 2, 3, 0);
        }
        chunks[0].emit_op_u16(Op::LOCAL_SET, 0, 0);
        chunks[0].emit_string_const(", ", 0);
        chunks[0].emit_op_u16(Op::LOCAL_SET, 4, 0);
        chunks[0].emit_string_const(": ", 0);
        chunks[0].emit_op_u16(Op::LOCAL_SET, 5, 0);
        if old {
            baseline::emit_render_separated(&mut chunks, 0, 0, 4, 5, 0);
        } else {
            json::emit_render_separated(&mut chunks, 0, 0, 4, 5, 0);
        }
        if call + 1 != calls {
            chunks[0].emit_op(Op::DROP, 0);
        }
    }
    chunks[0].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
fn array(values: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)))
}
fn object(entries: &[(&str, Value)]) -> Value {
    let mut object = Object::new();
    for (key, value) in entries {
        object.properties.insert((*key).into(), value.clone());
    }
    Value::Object(vybe_runtime::heap::alloc(object))
}
fn map(entries: Vec<(Value, Value)>) -> Value {
    let mut object = Object::new();
    object.kind = ObjectKind::Map(entries.into_iter().collect());
    Value::Object(vybe_runtime::heap::alloc(object))
}
fn vm(input: Value) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("input", input);
    vm
}
fn validate(chunks: &[Chunk]) {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(chunks))
        .unwrap();
}

#[test]
fn json_private_cursors_and_shared_bodies_preserve_nested_values_and_key_order() {
    let fixture = || {
        array(vec![
            Value::Null,
            Value::Bool(false),
            Value::I32(7),
            Value::String("é😀\n\"".into()),
            object(&[
                ("z", Value::I32(1)),
                ("a", array(vec![Value::String("nested".into())])),
            ]),
            map(vec![(Value::String("MapKey".into()), Value::I32(9))]),
        ])
    };
    for sort in [false, true] {
        for (own, prototype) in [(false, false), (true, false), (false, true), (true, true)] {
            let before = vm(fixture())
                .run(program(true, 2, false, sort, true, own, prototype))
                .unwrap();
            let chunks = program(false, 2, false, sort, true, own, prototype);
            validate(&chunks);
            let after = vm(fixture()).run(chunks).unwrap();
            assert_eq!(before, after);
        }
    }
    for old in [true, false] {
        assert_eq!(
            vm(array(vec![]))
                .run(program(old, 1, false, false, true, false, false))
                .unwrap(),
            Value::String("[]".into())
        );
        let ordered = object(&[
            ("b", Value::I32(2)),
            ("a", Value::I32(1)),
            (
                "__ps_key_order",
                array(vec![Value::String("b".into()), Value::String("a".into())]),
            ),
        ]);
        assert_eq!(
            vm(ordered)
                .run(program(old, 1, false, false, true, false, false))
                .unwrap(),
            Value::String("{\"b\": 2, \"a\": 1}".into())
        );
    }
}

#[test]
fn hidden_type_tags_preserve_existing_default_hook_gap() {
    use std::sync::{
        atomic::{AtomicUsize, Ordering},
        Arc,
    };
    for old in [true, false] {
        let input = array(vec![
            object(&[("__type", Value::String("Custom".into()))]),
            object(&[("__type", Value::String("Other".into()))]),
        ]);
        let mut runtime = vm(input);
        let count = Arc::new(AtomicUsize::new(0));
        let observed = count.clone();
        runtime.register_host_fn(
            "test:json",
            "hook",
            Box::new(move |_ctx, _args| {
                observed.fetch_add(1, Ordering::Relaxed);
                Value::String("encoded".into())
            }),
        );
        let chunks = program(old, 3, true, false, false, false, false);
        if !old {
            validate(&chunks);
        }
        assert_eq!(
            runtime.run(chunks).unwrap(),
            Value::String("[{}, {}]".into())
        );
        assert_eq!(count.load(Ordering::Relaxed), 0);
        let input = object(&[("__type", Value::String("Custom".into()))]);
        // hasOwn hides internal __type, so this baseline path neither calls
        // default nor throws for the tag. Keep that semantic gap explicit.
        assert_eq!(
            vm(input)
                .run(program(old, 1, false, false, false, false, false))
                .unwrap(),
            Value::String("{}".into())
        );
    }
}

#[test]
fn repeated_json_helpers_preserve_default_policy_and_finalize_idempotently() {
    for (own, prototype) in [(false, false), (true, false), (false, true), (true, true)] {
        let before = program(true, 20, false, false, true, own, prototype);
        let mut after = program(false, 20, false, false, true, own, prototype);
        let helpers: Vec<_> = after
            .iter()
            .filter(|c| c.name.starts_with("__json_"))
            .collect();
        assert_eq!(helpers.len(), 2);
        for helper in helpers {
            assert!(!helper.module_instance_fields_are_own_properties);
            assert!(!helper.module_members_on_prototype);
        }
        assert_eq!(
            before
                .iter()
                .filter(|c| c.name.starts_with("__json_"))
                .count(),
            40
        );
        validate(&after);
        let original = vybe_platform_wasm::write_wasm(&after);
        bundle::finalize_with_runtime_helpers(&mut after);
        assert_eq!(original, vybe_platform_wasm::write_wasm(&after));
        let size = |chunks: &[Chunk]| chunks.iter().map(|c| c.code.len()).sum::<usize>();
        assert!(size(&after) * 5 < size(&before));
        println!(
            "JSON calls=20 own={own} prototype={prototype} bytecode={}=>{} wasm={}=>{}",
            size(&before),
            size(&after),
            vybe_platform_wasm::write_wasm(&before).len(),
            original.len()
        );
    }
}

#[test]
#[ignore = "interleaved opt-level-0 diagnostic benchmark"]
fn benchmark_json_array_normalization_and_rendering() {
    use std::time::Instant;
    for n in [256, 2048] {
        let input = || array((0..n).map(|i| Value::I32(i)).collect());
        let before = program(true, 1, false, false, true, false, false);
        let after = program(false, 1, false, false, true, false, false);
        let mut old = vec![];
        let mut new = vec![];
        for sample in 0..5 {
            for previous in if sample % 2 == 0 {
                [true, false]
            } else {
                [false, true]
            } {
                let mut runtime = vm(input());
                let chunks = if previous {
                    before.clone()
                } else {
                    after.clone()
                };
                let start = Instant::now();
                runtime.run(chunks).unwrap();
                (if previous { &mut old } else { &mut new }).push(start.elapsed().as_micros());
            }
        }
        old.sort_unstable();
        new.sort_unstable();
        println!(
            "JSON items={n} samples=5 old_us={} new_us={} speedup={:.2}x",
            old[2],
            new[2],
            old[2] as f64 / new[2] as f64
        );
    }
}

#[test]
fn typed_and_prototype_protocol_hooks_preserve_helpers_existing_policy() {
    use std::sync::{
        atomic::{AtomicUsize, Ordering},
        Arc,
    };
    use vybe_runtime::value::Function;
    for (own, prototype) in [(false, false), (true, false), (false, true), (true, true)] {
        for inherited in [false, true] {
            for old in [true, false] {
                let mut chunks = program(old, 1, false, false, false, own, prototype);
                let role_index = chunks.len();
                let mut runtime = vm(Value::Null);
                let count = Arc::new(AtomicUsize::new(0));
                let observed = count.clone();
                runtime.register_host_fn(
                    "test:json",
                    "protocol",
                    Box::new(move |_ctx, _args| {
                        observed.fetch_add(1, Ordering::Relaxed);
                        Value::String("protocol-role".into())
                    }),
                );
                let mut role = Chunk::new("<json-protocol-role>");
                role.arity = 1;
                role.local_count = 1;
                role.emit_op_u16(Op::LOCAL_GET, 0, 0);
                let import = role.add_import("test:json", "protocol");
                role.emit_call(import, 1, 0);
                role.emit_op(Op::RETURN, 0);
                chunks.push(role);
                globals::declare_free_globals(&mut chunks);
                globals::normalize_global_table(&mut chunks);
                let mut function = Object::new();
                function.kind = ObjectKind::Function(Function {
                    name: Some("protocol-role".into()),
                    arity: 1,
                    // Host-created function values are not relocated by VM::run.
                    chunk_index: runtime.chunks.len() + role_index,
                    upvalues: vec![],
                });
                let function = Value::Object(vybe_runtime::heap::alloc(function));
                let key = vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Int);
                let mut ty = vybe_runtime::typedef::TypeDef::new("JsonTypedProtocol");
                ty.add_field(if inherited { "unrelated" } else { &key });
                let id = runtime.type_registry.register(ty);
                let mut input = Object::new_typed_with_fields(id, 1);
                input
                    .properties
                    .insert("__type".into(), Value::String("Custom".into()));
                if inherited {
                    input
                        .properties
                        .insert("__proto__".into(), object(&[(&key, function)]));
                } else {
                    input.fields[0] = function;
                }
                runtime.set_global("input", Value::Object(vybe_runtime::heap::alloc(input)));
                if !old {
                    validate(&chunks);
                }
                assert_eq!(
                    runtime.run(chunks).unwrap(),
                    if inherited {
                        // STRUCT_GET currently walks inherited accessors but
                        // ordinary-instance data properties on __proto__ are
                        // not walked by resolve_property. Preserve this gap.
                        Value::String("{}".into())
                    } else {
                        Value::String("\"protocol-role\"".into())
                    },
                    "old={old} own={own} prototype={prototype} inherited={inherited}"
                );
                assert_eq!(count.load(Ordering::Relaxed), if inherited { 0 } else { 1 });
            }
        }
    }
}

#[test]
fn same_named_source_functions_do_not_replace_generated_json_helpers() {
    use vybe_runtime::chunk::{LocalName, LocalOrigin};
    let mut source = vec![];
    for (name, arity, result) in [
        ("__json_normalize", 4, "user-normalize"),
        ("__json_render", 3, "user-render"),
    ] {
        let mut chunk = Chunk::new(name);
        chunk.arity = arity;
        chunk.local_count = arity as u16;
        // Even matching marker spelling is insufficient without Compiler origin.
        chunk.local_names.push(LocalName::new(
            0,
            format!("<json-helper:{name}:value>"),
            LocalOrigin::Source,
        ));
        chunk.emit_string_const(result, 0);
        chunk.emit_op(Op::RETURN, 0);
        source.push(chunk);
    }
    let mut chunks = program_seeded(false, 3, false, false, true, false, false, source);
    let original_user_bodies = [chunks[1].code.clone(), chunks[2].code.clone()];
    assert_eq!(
        chunks
            .iter()
            .filter(|c| c.name == "__json_normalize")
            .count(),
        2
    );
    assert_eq!(
        chunks.iter().filter(|c| c.name == "__json_render").count(),
        2
    );
    // Return JSON and the results of invoking both colliding source functions.
    let code_end = chunks[0].code.len() - 4;
    let line_end = chunks[0].lines.len() - 4;
    chunks[0].code.truncate(code_end);
    chunks[0].lines.truncate(line_end);
    for (index, arity) in [(1, 4), (2, 3)] {
        chunks[0].emit_op_u16(Op::REF_FUNC, index, 0);
        chunks[0].emit(0, 0);
        for _ in 0..arity {
            chunks[0].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        }
        chunks[0].emit_op_u8_u8(Op::CALL_REF, arity, 1, 0);
    }
    chunks[0].emit_array_new_fixed(0, 3, 0);
    chunks[0].emit_op(Op::RETURN, 0);
    validate(&chunks);
    let result = vm(array(vec![Value::I32(1), Value::I32(2)]))
        .run(chunks.clone())
        .unwrap();
    let Value::Object(result) = result else {
        panic!("expected result array");
    };
    let result = result.lock().unwrap();
    let ObjectKind::Array(items) = &result.kind else {
        panic!("expected result array");
    };
    assert_eq!(
        items,
        &vec![
            Value::String("[1, 2]".into()),
            Value::String("user-normalize".into()),
            Value::String("user-render".into())
        ]
    );
    assert_eq!(chunks[1].code, original_user_bodies[0]);
    assert_eq!(chunks[2].code, original_user_bodies[1]);
}
