use vybe_compiler::primitives::{bundle, functions, globals, platforms};
use vybe_runtime::value::{Function, Object, ObjectKind};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/text_optimization_format_map_hybrid_baseline.rs"]
mod baseline;
#[path = "fixtures/text_optimization_format_map_hybrid_candidate.rs"]
mod candidate;

fn program(old: bool, calls: usize, callback: bool) -> Vec<Chunk> {
    let mut entry = Chunk::new("<format-map-hybrid-optimization>");
    let body = if old {
        baseline::build_format_map(&mut entry)
    } else {
        candidate::build_format_map(&mut entry)
    };
    let mut chunks = vec![entry, body];
    if callback {
        let mut getitem = Chunk::new("<format-map-getitem>");
        getitem.arity = 2;
        getitem.local_count = 2;
        getitem.emit_op_u16(Op::LOCAL_GET, 1, 0);
        let index = getitem.add_import("test:format-map", "getitem");
        getitem.emit_call(index, 1, 0);
        getitem.emit_op(Op::RETURN, 0);
        chunks.push(getitem);
    }
    for call in 0..calls {
        chunks[0].emit_op_u16(Op::REF_FUNC, 1, 0);
        chunks[0].emit(0, 0);
        globals::emit_read(&mut chunks[0], "text", 0);
        globals::emit_read(&mut chunks[0], "dict", 0);
        functions::emit_call(&mut chunks[0], 2, 0);
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
fn dict() -> Value {
    let mut object = Object::new();
    for (key, value) in [
        ("a", Value::I32(17)),
        ("b", Value::Bool(true)),
        ("", Value::String("empty".into())),
        ("null", Value::Null),
        ("undef", Value::Undefined),
        ("é", Value::String("café".into())),
    ] {
        object.properties.insert(key.into(), value);
    }
    Value::Object(vybe_runtime::heap::alloc(object))
}
fn vm() -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm
}
fn result(
    runtime: &mut VM,
    chunks: &[Chunk],
    text: Value,
    dictionary: Value,
) -> Result<Value, String> {
    runtime.set_global("text", text);
    runtime.set_global("dict", dictionary);
    runtime.run(chunks.to_vec()).map_err(|mut e| {
        for frame in &mut e.call_stack {
            frame.offset = 0;
        }
        e.to_string()
    })
}
fn validate(chunks: &[Chunk]) {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(chunks))
        .unwrap();
}
#[test]
fn format_map_preserves_legacy_unicode_braces_coercion_and_errors() {
    let old = program(true, 1, false);
    let new = program(false, 1, false);
    validate(&old);
    validate(&new);
    let mut before = vm();
    let mut after = vm();
    for text in [
        "",
        "plain",
        "{a}:{b}:{a}",
        "{missing}:{null}:{undef}",
        "{}",
        "{{a}}",
        "}}x{{",
        "{a",
        "{",
        "{a.b}:{a:.2f}",
        "é😀{é}😀",
        "a\0{a}\n",
    ] {
        assert_eq!(
            result(&mut before, &old, Value::String(text.into()), dict()),
            result(&mut after, &new, Value::String(text.into()), dict()),
            "text={text:?}"
        );
    }
    assert_eq!(
        result(&mut after, &new, Value::String("{a}:{b}".into()), dict()).unwrap(),
        Value::String("17:true".into())
    );
    for text in [
        Value::Null,
        Value::Undefined,
        Value::I32(7),
        Value::Bool(false),
    ] {
        assert_eq!(
            result(&mut before, &old, text.clone(), dict()),
            result(&mut after, &new, text.clone(), dict()),
            "receiver={text:?}"
        );
    }
    for dictionary in [Value::Null, Value::Undefined, Value::I32(7)] {
        assert_eq!(
            result(
                &mut before,
                &old,
                Value::String("{a}".into()),
                dictionary.clone()
            ),
            result(&mut after, &new, Value::String("{a}".into()), dictionary)
        );
    }
    let size = |cs: &[Chunk]| cs.iter().map(|c| c.code.len()).sum::<usize>();
    for calls in [1, 20] {
        let old = program(true, calls, false);
        let new = program(false, calls, false);
        validate(&old);
        validate(&new);
        // Collection adds a few instructions; report size rather than require shrinkage.
        println!(
            "format_map calls={calls} bytecode={}=>{} wasm={}=>{}",
            size(&old),
            size(&new),
            vybe_platform_wasm::write_wasm(&old).len(),
            vybe_platform_wasm::write_wasm(&new).len()
        );
    }
}
#[test]
fn format_map_preserves_getitem_and_string_conversion_order() {
    use std::sync::{Arc, Mutex};
    for old in [true, false] {
        let mut runtime = vm();
        let chunks = program(old, 1, true);
        validate(&chunks);
        let log = Arc::new(Mutex::new(Vec::<String>::new()));
        let observed = log.clone();
        runtime.register_host_fn(
            "test:format-map",
            "getitem",
            Box::new(move |_ctx, args| {
                observed
                    .lock()
                    .unwrap()
                    .push(format!("get:{}", args[0].as_str()));
                Value::String(args[0].as_str().into())
            }),
        );
        let observed = log.clone();
        runtime.register_host_fn(
            "ecma:string",
            "String",
            Box::new(move |_ctx, args| {
                observed
                    .lock()
                    .unwrap()
                    .push(format!("str:{}", args[0].as_str()));
                Value::String(args[0].as_str().into())
            }),
        );
        let mut function = Object::new();
        function.kind = ObjectKind::Function(Function {
            name: Some("getitem".into()),
            arity: 2,
            chunk_index: runtime.chunks.len() + 2,
            upvalues: vec![],
        });
        let mut object = Object::new();
        object.properties.insert(
            "__getitem__".into(),
            Value::Object(vybe_runtime::heap::alloc(function)),
        );
        object.properties.insert("a".into(), Value::Null);
        object.properties.insert("b".into(), Value::Null);
        assert_eq!(
            result(
                &mut runtime,
                &chunks,
                Value::String("x{b}-{a}{b}y".into()),
                Value::Object(vybe_runtime::heap::alloc(object))
            )
            .unwrap(),
            Value::String("xb-aby".into())
        );
        assert_eq!(
            *log.lock().unwrap(),
            ["get:b", "str:b", "get:a", "str:a", "get:b", "str:b"]
        );
    }
}
#[test]
fn hybrid_boundaries_mixed_pieces_and_literals_match_typed_builder() {
    let old = program(true, 1, false);
    let new = program(false, 1, false);
    let mut before = vm();
    let mut after = vm();
    for (prefix, size, first_small) in [
        (124, 65535, false),
        (124, 65536, false),
        (125, 65535, false),
        (125, 65536, false),
        (0, 65536, false),
        (125, 65536, true),
    ] {
        let text = if first_small {
            format!("{}{b}{{a}}-{{a}}tail", "p".repeat(prefix), b = "{b}")
        } else {
            format!("{}{{a}}", "p".repeat(prefix))
        };
        let mut object = Object::new();
        object
            .properties
            .insert("a".into(), Value::String("x".repeat(size).into()));
        object
            .properties
            .insert("b".into(), Value::String("small".into()));
        let dictionary = Value::Object(vybe_runtime::heap::alloc(object));
        let expected = if first_small {
            format!(
                "{}small{}-{}tail",
                "p".repeat(prefix),
                "x".repeat(size),
                "x".repeat(size)
            )
        } else {
            format!("{}{}", "p".repeat(prefix), "x".repeat(size))
        };
        for (runtime, chunks) in [(&mut before, &old), (&mut after, &new)] {
            assert_eq!(
                result(
                    runtime,
                    chunks,
                    Value::String(text.clone().into()),
                    dictionary.clone()
                )
                .unwrap(),
                Value::String(expected.clone().into())
            );
        }
    }
    for value in ["é".repeat(65536), "😀".repeat(32768)] {
        let text = format!("{}{{a}}-{{a}}", "p".repeat(125));
        let expected = format!("{}{}-{}", "p".repeat(125), value, value);
        let mut object = Object::new();
        object
            .properties
            .insert("a".into(), Value::String(value.into()));
        let dictionary = Value::Object(vybe_runtime::heap::alloc(object));
        for (runtime, chunks) in [(&mut before, &old), (&mut after, &new)] {
            assert_eq!(
                result(
                    runtime,
                    chunks,
                    Value::String(text.clone().into()),
                    dictionary.clone()
                )
                .unwrap(),
                Value::String(expected.clone().into())
            );
        }
    }
}
#[test]
fn activated_hybrid_preserves_callback_and_conversion_order() {
    use std::sync::{Arc, Mutex};
    for old in [true, false] {
        let mut runtime = vm();
        let chunks = program(old, 1, true);
        validate(&chunks);
        let log = Arc::new(Mutex::new(Vec::<String>::new()));
        let seen = log.clone();
        runtime.register_host_fn(
            "test:format-map",
            "getitem",
            Box::new(move |_ctx, args| {
                let key = args[0].as_str();
                seen.lock().unwrap().push(format!("get:{key}"));
                Value::String(key.repeat(65536).into())
            }),
        );
        let seen = log.clone();
        runtime.register_host_fn(
            "ecma:string",
            "String",
            Box::new(move |_ctx, args| {
                let value = args[0].as_str();
                seen.lock()
                    .unwrap()
                    .push(format!("str:{}:{}", &value[..1], value.len()));
                Value::String(value.into())
            }),
        );
        let mut function = Object::new();
        function.kind = ObjectKind::Function(Function {
            name: Some("getitem".into()),
            arity: 2,
            chunk_index: runtime.chunks.len() + 2,
            upvalues: vec![],
        });
        let mut object = Object::new();
        object.properties.insert(
            "__getitem__".into(),
            Value::Object(vybe_runtime::heap::alloc(function)),
        );
        object.properties.insert("a".into(), Value::Null);
        object.properties.insert("b".into(), Value::Null);
        let text = format!("{}{{b}}-{{a}}{{b}}y", "p".repeat(125));
        let expected = format!(
            "{}{}-{}{}y",
            "p".repeat(125),
            "b".repeat(65536),
            "a".repeat(65536),
            "b".repeat(65536)
        );
        assert_eq!(
            result(
                &mut runtime,
                &chunks,
                Value::String(text.into()),
                Value::Object(vybe_runtime::heap::alloc(object))
            )
            .unwrap(),
            Value::String(expected.into())
        );
        assert_eq!(
            *log.lock().unwrap(),
            [
                "get:b",
                "str:b:65536",
                "get:a",
                "str:a:65536",
                "get:b",
                "str:b:65536"
            ]
        );
    }
}
#[test]
#[ignore = "interleaved opt-level-0 fragment accumulation experiment"]
fn benchmark_format_map_hybrid() {
    use std::time::Instant;
    let old = program(true, 1, false);
    let new = program(false, 1, false);
    for (label, text, value) in [
        ("literal96", "x".repeat(96), "17".to_owned()),
        ("literal2048", "x".repeat(2048), "17".to_owned()),
        ("placeholders256", "{a}".repeat(256), "17".to_owned()),
        ("placeholders2048", "{a}".repeat(2048), "17".to_owned()),
        ("output12k", "{a}{a}{a}".to_owned(), "x".repeat(4096)),
        ("output192k", "{a}{a}{a}".to_owned(), "x".repeat(65536)),
        ("pieces64x64k", "{a}".repeat(64), "x".repeat(65536)),
        ("pieces256x64k", "{a}".repeat(256), "x".repeat(65536)),
    ] {
        let mut before = vm();
        let mut after = vm();
        let mut dictionary = Object::new();
        dictionary
            .properties
            .insert("a".into(), Value::String(value.into()));
        let dictionary = Value::Object(vybe_runtime::heap::alloc(dictionary));
        let expected = result(
            &mut before,
            &old,
            Value::String(text.clone().into()),
            dictionary.clone(),
        )
        .unwrap();
        let mut old_times = vec![];
        let mut new_times = vec![];
        for sample in 0..7 {
            for previous in if sample % 2 == 0 {
                [true, false]
            } else {
                [false, true]
            } {
                let (runtime, chunks, times) = if previous {
                    (&mut before, &old, &mut old_times)
                } else {
                    (&mut after, &new, &mut new_times)
                };
                runtime.set_global("text", Value::String(text.clone().into()));
                runtime.set_global("dict", dictionary.clone());
                let prepared = chunks.clone();
                let started = Instant::now();
                let actual = runtime.run(prepared).unwrap();
                let elapsed = started.elapsed();
                assert_eq!(actual, expected);
                times.push(elapsed.as_nanos());
            }
        }
        old_times.sort();
        new_times.sort();
        println!(
            "format_map hybrid {label} old={}ns new={}ns gain={:.2}x",
            old_times[3],
            new_times[3],
            old_times[3] as f64 / new_times[3] as f64
        );
    }
}
