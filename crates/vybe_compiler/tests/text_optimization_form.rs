use vybe_compiler::primitives::{bundle, globals, http_form, platforms};
use vybe_runtime::value::ObjectKind;
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/text_optimization_form_baseline.rs"]
mod baseline;
fn program(old: bool, calls: usize) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<text-form-optimization>")];
    for call in 0..calls {
        globals::emit_read(&mut chunks[0], "body", 0);
        if old {
            baseline::emit_parse_urlencoded(&mut chunks, 0, 0);
        } else {
            http_form::emit_parse_urlencoded(&mut chunks, 0, 0);
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
fn vm(body: Value) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("body", body);
    vm
}
fn entries(value: Value) -> Vec<(String, String)> {
    let Value::Object(object) = value else {
        panic!("expected form map")
    };
    let object = object.lock().unwrap();
    let ObjectKind::Map(map) = &object.kind else {
        panic!("expected form map")
    };
    map.iter()
        .map(|(k, v)| (k.as_str().to_owned(), v.as_str().to_owned()))
        .collect()
}
fn result(old: bool, body: Value) -> Result<Vec<(String, String)>, String> {
    vm(body)
        .run(program(old, 1))
        .map(entries)
        .map_err(|mut error| {
            // The emitted body shrinks, so instruction offsets differ. Keep
            // error type/message, frame names/lines and diagnostics intact.
            for frame in &mut error.call_stack {
                frame.offset = 0;
            }
            error.to_string()
        })
}
#[test]
fn separator_reuse_preserves_decoding_duplicates_unicode_and_errors() {
    for body in [
        Value::Null,
        Value::Undefined,
        Value::String("".into()),
        Value::String("a=1&b=2".into()),
        Value::String("name=John+Smith&user%20name=caf%C3%A9".into()),
        Value::String("a=first&a=second&bare&&=skip&empty=".into()),
        Value::String("😀é=one=two&literal=%2B%26%3D".into()),
        Value::String("a=%ZZ".into()),
        Value::String("a=%".into()),
        Value::String("a=%C3".into()),
        Value::String("a=%ED%A0%80".into()),
    ] {
        assert_eq!(
            result(true, body.clone()),
            result(false, body.clone()),
            "body={body:?}"
        );
    }
    assert_eq!(
        result(false, Value::String("a=first&a=last&bare&empty=".into())).unwrap(),
        vec![("a".into(), "last".into()), ("empty".into(), "".into())]
    );
    assert_eq!(
        result(false, Value::String("😀é=one=two&literal=%2B%26%3D".into())).unwrap(),
        vec![
            ("😀é".into(), "one=two".into()),
            ("literal".into(), "+&=".into())
        ]
    );
    for old in [true, false] {
        for calls in [1, 20] {
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&vybe_platform_wasm::write_wasm(&program(old, calls)))
                .unwrap();
        }
    }
    let old = program(true, 20);
    let new = program(false, 20);
    let size = |chunks: &[Chunk]| chunks.iter().map(|c| c.code.len()).sum::<usize>();
    assert!(size(&new) < size(&old));
    println!(
        "URL form calls=20 bytecode={}=>{} wasm={}=>{}",
        size(&old),
        size(&new),
        vybe_platform_wasm::write_wasm(&old).len(),
        vybe_platform_wasm::write_wasm(&new).len()
    );
}
#[test]
fn private_pair_separator_is_searched_once_per_segment() {
    use std::sync::{
        atomic::{AtomicUsize, Ordering},
        Arc,
    };
    for old in [true, false] {
        let mut runtime = vm(Value::String("a=1&bare&&b=two=three&c=".into()));
        let count = Arc::new(AtomicUsize::new(0));
        let observed = count.clone();
        runtime.register_host_fn(
            "ecma:string",
            "indexOf",
            Box::new(move |_ctx, args| {
                observed.fetch_add(1, Ordering::Relaxed);
                let source: Vec<_> = args[0].as_str().encode_utf16().collect();
                let needle: Vec<_> = args[1].as_str().encode_utf16().collect();
                let found = source
                    .windows(needle.len())
                    .position(|w| w == needle)
                    .map(|p| p as i32)
                    .unwrap_or(-1);
                Value::I32(found)
            }),
        );
        let output = entries(runtime.run(program(old, 1)).unwrap());
        assert_eq!(
            output,
            vec![
                ("a".into(), "1".into()),
                ("b".into(), "two=three".into()),
                ("c".into(), "".into())
            ]
        );
        assert_eq!(count.load(Ordering::Relaxed), if old { 8 } else { 5 });
    }
}
#[test]
#[ignore = "interleaved opt-level-0 diagnostic benchmark"]
fn benchmark_urlencoded_separator_reuse() {
    use std::time::Instant;
    for n in [16, 256, 2048] {
        let body = (0..n)
            .map(|i| format!("field{i}=some+value%3D{i}"))
            .collect::<Vec<_>>()
            .join("&");
        let before = program(true, 1);
        let after = program(false, 1);
        let mut old = vec![];
        let mut new = vec![];
        for sample in 0..5 {
            for previous in if sample % 2 == 0 {
                [true, false]
            } else {
                [false, true]
            } {
                let mut runtime = vm(Value::String(body.clone().into()));
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
            "URL form fields={n} samples=5 old_us={} new_us={} speedup={:.2}x",
            old[2],
            new[2],
            old[2] as f64 / new[2] as f64
        );
    }
}
