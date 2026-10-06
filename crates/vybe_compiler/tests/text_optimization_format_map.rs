use vybe_compiler::primitives::{bundle, functions, globals, platforms, strings};
use vybe_runtime::value::{Function, Object, ObjectKind};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/text_optimization_format_map_baseline.rs"]
mod baseline;

fn program(old: bool, calls: usize, callback: bool) -> Vec<Chunk> {
    let mut entry = Chunk::new("<format-map-optimization>");
    let body = if old {
        baseline::build_format_map(&mut entry)
    } else {
        strings::build_format_map(&mut entry)
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
        assert!(size(&new) < size(&old));
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
fn private_operations_remove_emitted_dynamic_type_probes() {
    fn calls_to(chunk: &Chunk, module: &str, name: &str) -> usize {
        let mut offset = 0;
        let mut calls = 0;
        while offset < chunk.code.len() {
            let op = Op::from_code_at(&chunk.code, offset);
            if op == Op::CALL {
                let index =
                    u16::from_be_bytes([chunk.code[offset + 4], chunk.code[offset + 5]]) as usize;
                let import = &chunk.imports[index];
                if import.module == module && import.name == name {
                    calls += 1;
                }
            }
            offset += 4 + op.operand_format().size_in(&chunk.code, offset + 4);
        }
        calls
    }
    for old in [true, false] {
        let chunks = program(old, 1, false);
        // Count actual emitted call instructions, including untaken branches.
        // VM intrinsics bypass host overrides, so these are static counts.
        assert_eq!(
            calls_to(&chunks[1], "wasm:js-string", "test"),
            if old { 12 } else { 0 }
        );
        assert_eq!(
            calls_to(&chunks[1], "wasm:js-number", "test"),
            if old { 4 } else { 0 }
        );
        assert_eq!(
            result(&mut vm(), &chunks, Value::String("x{a}y".into()), dict()).unwrap(),
            Value::String("x17y".into())
        );
    }
}
#[test]
#[ignore = "interleaved opt-level-0 diagnostic benchmark"]
fn benchmark_format_map_private_operations() {
    use std::time::Instant;
    let old = program(true, 1, false);
    let new = program(false, 1, false);
    for text in [
        "literal".repeat(16),
        "x{a}:{b}".repeat(32),
        "x{a}:{b}".repeat(256),
    ] {
        let mut before = vm();
        let mut after = vm();
        let mut old_times = vec![];
        let mut new_times = vec![];
        let expected = result(
            &mut before,
            &old,
            Value::String(text.clone().into()),
            dict(),
        )
        .unwrap();
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
                runtime.set_global("dict", dict());
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
            "format_map utf16={} old={}ns new={}ns gain={:.2}x",
            text.encode_utf16().count(),
            old_times[3],
            new_times[3],
            old_times[3] as f64 / new_times[3] as f64
        );
    }
}
