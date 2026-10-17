//! Frozen comparison of actual production snapshot-set emitters.
use std::sync::{Arc, Mutex};
use vybe_compiler::primitives::{bundle, globals, platforms, sets};
use vybe_runtime::value::{ContinuationPhase, ContinuationState, Function, Object, ObjectKind};
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/collection_optimization_sets_snapshot.rs"]
mod baseline;
fn object(object: Object) -> Value {
    Value::Object(vybe_runtime::heap::alloc(object))
}
fn array(values: Vec<Value>) -> Value {
    object(Object::new_array(values))
}
fn function(index: usize, arity: u8) -> Value {
    object(Object {
        descriptor: None,
            kind: ObjectKind::Function(Function {
            name: Some("callback".into()),
            arity,
            chunk_index: index,
            upvalues: vec![],
        }),
        ..Object::new()
    })
}
fn set(values: Vec<Value>) -> Value {
    let mut o = Object::new();
    o.kind = ObjectKind::Set(values.into_iter().collect());
    object(o)
}
fn vm(values: &[(&str, Value)]) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    for (k, v) in values {
        vm.set_global(*k, v.clone());
    }
    vm
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
fn error(mut e: vybe_runtime::VMError) -> String {
    for f in &mut e.call_stack {
        f.offset = 0;
    }
    format!("{e:?}")
}
fn program(old: bool, resync: bool, callbacks: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("<sets-private-snapshot>");
    globals::emit_read(&mut root, "input", 0);
    let mut chunks = vec![root];
    if callbacks {
        let mut render = Chunk::new("<render-callback>");
        render.arity = 1;
        render.local_count = 1;
        render.emit_string_const("render", 0);
        render.emit_i32_const(0, 0);
        let observe = render.add_import("ecma:string", "charAt");
        render.emit_call(observe, 2, 0);
        render.emit_op(Op::RETURN, 0);
        chunks.push(render);
        let mut iterator = Chunk::new("<custom-iterator>");
        globals::emit_read(&mut iterator, "continuation", 0);
        iterator.emit_op(Op::RETURN, 0);
        chunks.push(iterator);
        let mut getitem = Chunk::new("<custom-getitem>");
        getitem.arity = 2;
        getitem.local_count = 2;
        getitem.emit_string_const("read", 0);
        getitem.emit_op_u16(Op::LOCAL_GET, 1, 0);
        let observe = getitem.add_import("ecma:string", "charAt");
        getitem.emit_call(observe, 2, 0);
        getitem.emit_op(Op::RETURN, 0);
        chunks.push(getitem);
    }
    match (old, resync) {
        (true, true) => baseline::emit_resync_snapshot_keys(&mut chunks, 0, 0),
        (true, false) => baseline::emit_from_iterable_snapshot(&mut chunks, 0, 0),
        (false, true) => sets::emit_resync_snapshot_keys(&mut chunks, 0, 0),
        (false, false) => sets::emit_from_iterable_snapshot(&mut chunks, 0, 0),
    }
    finish(chunks)
}
fn keys(value: &Value) -> Vec<(Value, Value)> {
    let Value::Object(o) = value else {
        panic!("set/object expected")
    };
    let keys = o
        .lock()
        .unwrap()
        .properties
        .get("__snapshot_keys")
        .cloned()
        .unwrap_or(Value::Null);
    let Value::Object(keys) = keys else {
        return vec![];
    };
    let keys = keys.lock().unwrap();
    let ObjectKind::Map(m) = &keys.kind else {
        panic!("map expected")
    };
    m.iter().map(|(k, v)| (k.clone(), v.clone())).collect()
}
fn key_labels(value: &Value) -> Vec<(String, String)> {
    keys(value)
        .iter()
        .map(|(k, v)| (format!("{k:?}"), format!("{v:?}")))
        .collect()
}
#[test]
fn scalar_snapshot_paths_keep_key_order_dedupe_and_types() {
    for resync in [false, true] {
        let programs = [program(true, resync, false), program(false, resync, false)];
        for chunks in &programs {
            validate(chunks);
        }
        for values in [
            vec![],
            vec![
                Value::I32(1),
                Value::I32(1),
                Value::Bool(false),
                Value::Null,
                Value::String("é😀".into()),
            ],
        ] {
            let results = programs.each_ref().map(|chunks| {
                let input = if resync {
                    set(values.clone())
                } else {
                    array(values.clone())
                };
                vm(&[("input", input)])
                    .run(chunks.clone())
                    .map(|v| key_labels(&v))
                    .map_err(error)
            });
            assert_eq!(results[0], results[1]);
        }
    }
}
#[test]
fn real_to_json_callbacks_cannot_reach_or_rewrite_private_materialization() {
    let programs = [program(true, true, true), program(false, true, true)];
    for chunks in &programs {
        validate(chunks);
    }
    let outcomes = programs.each_ref().map(|chunks| {
        let elements: Vec<Value> = (0..3)
            .map(|id| {
                let mut o = Object::new();
                o.properties.insert("id".into(), Value::I32(id));
                o.properties.insert(
                    vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Eq).to_string(),
                    function(1, 1),
                );
                o.properties.insert("toJSON".into(), function(1, 1));
                object(o)
            })
            .collect();
        let input = set(elements.clone());
        let observed_input = input.clone();
        let observed_elements = elements.clone();
        let calls = Arc::new(Mutex::new(Vec::new()));
        let observed = calls.clone();
        let mut runtime = vm(&[("input", input.clone())]);
        runtime.register_host_fn(
            "ecma:string",
            "charAt",
            Box::new(move |ctx, args| {
                assert_eq!(args[0], Value::String("render".into()));
                let this = ctx.current_js_this();
                let Value::Object(this) = this else {
                    panic!("real toJSON receiver required")
                };
                let id = this.lock().unwrap().properties["id"].as_i32();
                observed.lock().unwrap().push(id);
                let Value::Object(set) = &observed_input else {
                    unreachable!()
                };
                let mut set = set.lock().unwrap();
                let ObjectKind::Set(values) = &mut set.kind else {
                    unreachable!()
                };
                values.shift_remove(&observed_elements[2]);
                values.insert(Value::I32(99));
                drop(set);
                Value::String(format!("key{id}").into())
            }),
        );
        let result = runtime.run(chunks.clone()).map_err(error);
        let result = result.map(|v| {
            let keys = keys(&v);
            assert_eq!(keys.len(), 3);
            for (index, (_, value)) in keys.iter().enumerate() {
                assert_eq!(value, &elements[index]);
            }
            keys.iter()
                .map(|(key, _)| format!("{key:?}"))
                .collect::<Vec<_>>()
        });
        let trace = calls.lock().unwrap().clone();
        (result, trace)
    });
    assert_eq!(outcomes[0], outcomes[1]);
    assert_eq!(outcomes[0].1, vec![0, 1, 2]);
}
#[test]
fn actual_custom_iterable_continuation_keeps_both_reads() {
    let programs = [program(true, true, true), program(false, true, true)];
    for chunks in &programs {
        validate(chunks);
    }
    let outcomes = programs.each_ref().map(|chunks| {
        let mut cont = Object::new();
        cont.kind = ObjectKind::Continuation(ContinuationState {
            entry: Value::Null,
            saved: Box::new(Mutex::new(None)),
            state: Mutex::new(ContinuationPhase::Ready),
        });
        cont.properties.insert("length".into(), Value::I32(2));
        for index in 0..2 {
            cont.properties.insert(index.to_string(), Value::Null);
        }
        cont.properties.insert("__getitem__".into(), function(3, 2));
        let continuation = object(cont);
        let mut source = Object::new();
        source.properties.insert("iterator".into(), function(2, 0));
        let source = object(source);
        let calls = Arc::new(Mutex::new(Vec::new()));
        let observed = calls.clone();
        let mut runtime = vm(&[("input", source), ("continuation", continuation)]);
        runtime.register_host_fn(
            "ecma:string",
            "charAt",
            Box::new(move |_, args| {
                assert_eq!(args[0], Value::String("read".into()));
                let mut trace = observed.lock().unwrap();
                trace.push(args[1].as_i32());
                Value::String(format!("read{}", trace.len()).into())
            }),
        );
        let result = runtime
            .run(chunks.clone())
            .map(|v| key_labels(&v))
            .map_err(error);
        let trace = calls.lock().unwrap().clone();
        (result, trace)
    });
    assert_eq!(outcomes[0], outcomes[1]);
    assert_eq!(outcomes[0].1, vec![0, 0, 1, 1]);
    assert_eq!(outcomes[0].0.as_ref().unwrap().len(), 2);
}
#[test]
fn render_error_preserves_callback_count_and_partial_sidecar_mutation() {
    let outcomes = [program(true, true, true), program(false, true, true)].map(|chunks| {
        let mut element = Object::new();
        element.properties.insert(
            vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Eq).to_string(),
            function(1, 1),
        );
        element.properties.insert("toJSON".into(), function(1, 1));
        let input = set(vec![object(element)]);
        let calls = Arc::new(Mutex::new(0));
        let observed = calls.clone();
        let mut runtime = vm(&[("input", input.clone())]);
        runtime.register_host_fn(
            "ecma:string",
            "charAt",
            Box::new(move |ctx, _| {
                *observed.lock().unwrap() += 1;
                ctx.throw_value(Value::String("render-error".into()));
                Value::Undefined
            }),
        );
        let result = runtime.run(chunks).map(|v| key_labels(&v)).map_err(error);
        let count = *calls.lock().unwrap();
        (result, keys(&input).len(), count)
    });
    assert_eq!(outcomes[0], outcomes[1]);
    assert_eq!(outcomes[0].2, 1);
    assert!(outcomes[0].0.is_err());
}
#[test]
#[ignore = "interleaved opt-level-0 debug diagnostic"]
fn benchmark_private_snapshot_paths() {
    for resync in [false, true] {
        let programs = [program(true, resync, false), program(false, resync, false)];
        for count in [256, 8192] {
            let mut times = [vec![], vec![]];
            for sample in 0..9 {
                for which in [sample % 2, 1 - sample % 2] {
                    let values = (0..count).map(Value::I32).collect();
                    let input = if resync { set(values) } else { array(values) };
                    let mut runtime = vm(&[("input", input)]);
                    let chunks = programs[which].clone();
                    let start = std::time::Instant::now();
                    let result = runtime.run(chunks).unwrap();
                    times[which].push(start.elapsed().as_micros());
                    assert_eq!(keys(&result).len(), count as usize);
                }
            }
            for time in &mut times {
                time.sort_unstable();
            }
            eprintln!(
                "sets-snapshot resync={resync} count={count} before={}us after={}us ratio={:.2}x",
                times[0][4],
                times[1][4],
                times[0][4] as f64 / times[1][4] as f64
            );
        }
    }
}

#[test]
fn snapshot_private_arithmetic_reduces_full_module_emission() {
    for resync in [false, true] {
        let old = program(true, resync, false);
        let new = program(false, resync, false);
        validate(&old);
        validate(&new);
        let bytes = |chunks: &[Chunk]| chunks.iter().map(|c| c.code.len()).sum::<usize>();
        assert!(bytes(&new) < bytes(&old));
        eprintln!(
            "sets-snapshot resync={resync} bytecode={}/{} wasm={}/{}",
            bytes(&old),
            bytes(&new),
            vybe_platform_wasm::write_wasm(&old).len(),
            vybe_platform_wasm::write_wasm(&new).len()
        );
    }
}

#[test]
fn constructor_render_callbacks_keep_live_source_reads_and_error_order() {
    for fail in [false,true] {
        let outcomes=[program(true,false,true),program(false,false,true)].map(|chunks| {
            validate(&chunks);
            let elements=(0..3).map(|id| {
                let mut o=Object::new();
                o.properties.insert("id".into(),Value::I32(id));
                o.properties.insert(vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Eq).to_string(),function(1,1));
                o.properties.insert("toJSON".into(),function(1,1));
                object(o)
            }).collect();
            let input=array(elements);let observed_input=input.clone();
            let calls=Arc::new(Mutex::new(Vec::new()));let observed=calls.clone();
            let mut runtime=vm(&[("input",input)]);
            runtime.register_host_fn("ecma:string","charAt",Box::new(move |ctx,args| {
                assert_eq!(args[0],Value::String("render".into()));
                let Value::Object(this)=ctx.current_js_this() else {panic!("render receiver")};
                let id=this.lock().unwrap().properties["id"].as_i32();
                observed.lock().unwrap().push(id);
                if id==0 {
                    let Value::Object(input)=&observed_input else {unreachable!()};
                    let mut input=input.lock().unwrap();
                    let ObjectKind::Array(values)=&mut input.kind else {unreachable!()};
                    values[1]=Value::String("changed".into());
                }
                if fail && id==2 {ctx.throw_value(Value::String("constructor-render-error".into()));return Value::Undefined;}
                Value::String(format!("key{id}").into())
            }));
            let result=runtime.run(chunks).map(|v|key_labels(&v)).map_err(error);
            let trace=calls.lock().unwrap().clone();(result,trace)
        });
        assert_eq!(outcomes[0],outcomes[1]);
        assert_eq!(outcomes[0].1,vec![0,2]);
        if fail {assert!(outcomes[0].0.is_err());} else {assert_eq!(outcomes[0].0.as_ref().unwrap().len(),3);}
    }
}

