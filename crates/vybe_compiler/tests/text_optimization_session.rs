use vybe_compiler::primitives::{bundle, globals, http_session, platforms};
use vybe_runtime::value::Object;
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/text_optimization_session_byte_cache_baseline.rs"]
mod baseline;
fn program(old: bool, calls: usize, collect: bool) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<session-id-optimization>")];
    for call in 0..calls {
        if old {
            baseline::emit_new_id(&mut chunks, 0, 0);
        } else {
            http_session::emit_new_id(&mut chunks, 0, 0);
        }
        if !collect && call + 1 != calls {
            chunks[0].emit_op(Op::DROP, 0);
        }
    }
    if collect {
        chunks[0].emit_array_new_fixed(
            0,
            u16::try_from(calls).expect("fixture call count fits u16"),
            0,
        );
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
fn vm() -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm
}
fn validate(chunks: &[Chunk]) {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(chunks))
        .unwrap();
}
fn array_reads(chunk: &Chunk) -> usize {
    let mut offset = 0;
    let mut reads = 0;
    while offset < chunk.code.len() {
        let op = Op::from_code_at(&chunk.code, offset);
        if op == Op::ARRAY_GET {
            reads += 1;
        }
        offset += 4 + op.operand_format().size_in(&chunk.code, offset + 4);
    }
    reads
}

#[test]
fn cached_byte_preserves_every_byte_and_empty_buffers() {
    use std::sync::{
        atomic::{AtomicUsize, Ordering},
        Arc,
    };
    for bytes in [
        vec![],
        vec![0, 1, 15, 16, 127, 128, 254, 255],
        (0..=255).collect::<Vec<i32>>(),
    ] {
        let expected = bytes.iter().map(|b| format!("{b:02x}")).collect::<String>();
        for old in [true, false] {
            let mut runtime = vm();
            let reads = Arc::new(AtomicUsize::new(0));
            let observed = reads.clone();
            let bytes = bytes.clone();
            runtime.register_host_fn(
                "wasi:random/random",
                "get-random-bytes",
                Box::new(move |_ctx, args| {
                    assert_eq!(args[0].as_i32(), 16);
                    observed.fetch_add(1, Ordering::Relaxed);
                    array(bytes.iter().copied().map(Value::I32).collect())
                }),
            );
            let chunks = program(old, 1, false);
            validate(&chunks);
            assert_eq!(
                runtime.run(chunks).unwrap(),
                Value::String(expected.clone().into())
            );
            assert_eq!(reads.load(Ordering::Relaxed), 1);
        }
    }
    for calls in [1, 20] {
        let old = program(true, calls, false);
        let new = program(false, calls, false);
        validate(&new);
        assert_eq!(array_reads(&old[0]), calls * 2);
        assert_eq!(array_reads(&new[0]), calls);
        let size = |chunks: &[Chunk]| chunks.iter().map(|c| c.code.len()).sum::<usize>();
        assert!(size(&new) < size(&old));
        println!(
            "session IDs calls={calls} bytecode={}=>{} wasm={}=>{}",
            size(&old),
            size(&new),
            vybe_platform_wasm::write_wasm(&old).len(),
            vybe_platform_wasm::write_wasm(&new).len()
        );
    }
}
#[test]
fn every_id_requests_fresh_entropy_and_preserves_nibble_substrings() {
    use std::sync::{
        atomic::{AtomicUsize, Ordering},
        Arc,
    };
    use vybe_runtime::value::ObjectKind;
    for old in [true, false] {
        let mut runtime = vm();
        let requests = Arc::new(AtomicUsize::new(0));
        let observed = requests.clone();
        runtime.register_host_fn(
            "wasi:random/random",
            "get-random-bytes",
            Box::new(move |_ctx, args| {
                assert_eq!(args[0].as_i32(), 16);
                let index = observed.fetch_add(1, Ordering::Relaxed);
                array((0..16).map(|b| Value::I32((b + index) as i32)).collect())
            }),
        );
        let slices = Arc::new(AtomicUsize::new(0));
        let observed = slices.clone();
        runtime.register_host_fn(
            "wasm:js-string",
            "substring",
            Box::new(move |_ctx, args| {
                observed.fetch_add(1, Ordering::Relaxed);
                let start = args[1].as_i32() as usize;
                let end = args[2].as_i32() as usize;
                Value::String(args[0].as_str()[start..end].into())
            }),
        );
        let output = runtime.run(program(old, 3, true)).unwrap();
        let Value::Object(output) = output else {
            panic!("expected IDs")
        };
        let output = output.lock().unwrap();
        let ObjectKind::Array(items) = &output.kind else {
            panic!("expected IDs")
        };
        let expected: Vec<_> = (0..3)
            .map(|offset| {
                Value::String(
                    (0..16)
                        .map(|b| format!("{:02x}", b + offset))
                        .collect::<String>()
                        .into(),
                )
            })
            .collect();
        assert_eq!(items, &expected);
        assert_eq!(requests.load(Ordering::Relaxed), 3);
        assert_eq!(slices.load(Ordering::Relaxed), 96);
    }
}
#[test]
#[ignore = "interleaved opt-level-0 diagnostic benchmark with deterministic entropy"]
fn benchmark_session_cached_byte_emission() {
    use std::time::Instant;
    for calls in [1, 20, 100] {
        let before = program(true, calls, false);
        let after = program(false, calls, false);
        let mut old = vec![];
        let mut new = vec![];
        for sample in 0..7 {
            for previous in if sample % 2 == 0 {
                [true, false]
            } else {
                [false, true]
            } {
                let mut runtime = vm();
                runtime.register_host_fn(
                    "wasi:random/random",
                    "get-random-bytes",
                    Box::new(|_ctx, args| {
                        assert_eq!(args[0].as_i32(), 16);
                        array((0..16).map(Value::I32).collect())
                    }),
                );
                let chunks = if previous {
                    before.clone()
                } else {
                    after.clone()
                };
                let start = Instant::now();
                assert_eq!(
                    runtime.run(chunks).unwrap(),
                    Value::String("000102030405060708090a0b0c0d0e0f".into())
                );
                (if previous { &mut old } else { &mut new }).push(start.elapsed().as_micros());
            }
        }
        old.sort_unstable();
        new.sort_unstable();
        println!("session IDs calls={calls} samples=7 deterministic_entropy old_us={} new_us={} speedup={:.2}x",old[3],new[3],old[3] as f64/new[3] as f64);
    }
}
