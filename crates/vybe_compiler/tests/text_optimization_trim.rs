use vybe_compiler::primitives::{bundle, globals, platforms, strings};
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/text_optimization_trim_baseline.rs"]
mod baseline;
const DEFAULT: &str = " \t\n\r\0\u{b}";
fn program(old: bool, explicit: bool, opts: strings::TrimOptions, calls: usize) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<trim-optimization>")];
    for call in 0..calls {
        globals::emit_read(&mut chunks[0], "text", 0);
        if explicit {
            globals::emit_read(&mut chunks[0], "chars", 0);
        }
        if old {
            baseline::emit_trim_chars(&mut chunks, 0, if explicit { 2 } else { 1 }, opts, 0);
        } else {
            strings::emit_trim_chars(&mut chunks, 0, if explicit { 2 } else { 1 }, opts, 0);
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
fn vm(text: Value, chars: Value) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("text", text);
    vm.set_global("chars", chars);
    vm
}
fn run(
    old: bool,
    text: Value,
    chars: Value,
    explicit: bool,
    opts: strings::TrimOptions,
) -> Result<Value, String> {
    let mut runtime = vm(Value::Null, Value::Null);
    run_cached(&mut runtime, &program(old, explicit, opts, 1), text, chars)
}

fn run_cached(
    runtime: &mut VM,
    chunks: &[Chunk],
    text: Value,
    chars: Value,
) -> Result<Value, String> {
    runtime.set_global("text", text);
    runtime.set_global("chars", chars);
    // VM::run resets execution state and relocates each incoming chunk copy.
    runtime.run(chunks.to_vec()).map_err(|mut error| {
        for frame in &mut error.call_stack {
            frame.offset = 0;
        }
        error.to_string()
    })
}
#[test]
fn static_ascii_membership_preserves_utf16_boundaries_and_explicit_sets() {
    for defaults in [
        Some(DEFAULT),
        Some(""),
        Some("aabbxyz"),
        Some("\u{fffd}"),
        Some("é😀"),
        None,
    ] {
        for (left, right) in [(true, true), (true, false), (false, true), (false, false)] {
            let opts = strings::TrimOptions {
                left,
                right,
                default_chars: defaults,
            };
            for explicit in [false, true] {
                println!("trim matrix defaults={defaults:?} left={left} right={right} explicit={explicit}");
                let before = program(true, explicit, opts, 1);
                let after = program(false, explicit, opts, 1);
                let mut old_vm = vm(Value::Null, Value::Null);
                let mut new_vm = vm(Value::Null, Value::Null);
                for text in [
                    "",
                    " abc ",
                    "\0\t\nhello\r\u{b}",
                    "aaaa",
                    "xyzxyz",
                    "éabcé",
                    " 😀é\u{fffd} ",
                    "😀x😀",
                    "e\u{301}",
                    "\u{fffd}x\u{fffd}",
                ] {
                    for chars in [DEFAULT, "", "axyz", "\u{fffd}", "é😀", "[]$.*"] {
                        assert_eq!(run_cached(&mut old_vm, &before, Value::String(text.into()), Value::String(chars.into())),run_cached(&mut new_vm, &after, Value::String(text.into()), Value::String(chars.into())),"text={text:?} chars={chars:?} explicit={explicit} defaults={defaults:?} left={left} right={right}");
                    }
                }
            }
        }
    }
    let ascii_opts = strings::TrimOptions::both(Some(DEFAULT));
    let before_ascii = program(true, false, ascii_opts, 1);
    let after_ascii = program(false, false, ascii_opts, 1);
    let mut old_ascii_vm = vm(Value::Null, Value::Null);
    let mut new_ascii_vm = vm(Value::Null, Value::Null);
    for cp in 0..128 {
        let text = format!(
            "{}hello{}",
            char::from_u32(cp).unwrap(),
            char::from_u32(cp).unwrap()
        );
        assert_eq!(
            run_cached(
                &mut old_ascii_vm,
                &before_ascii,
                Value::String(text.clone().into()),
                Value::Null
            ),
            run_cached(
                &mut new_ascii_vm,
                &after_ascii,
                Value::String(text.into()),
                Value::Null
            ),
            "ASCII cp={cp}"
        );
    }
    let opts = strings::TrimOptions::both(Some(DEFAULT));
    assert_eq!(
        run(
            false,
            Value::String("\0\t\n😀é x😀\r\u{b} ".into()),
            Value::Null,
            false,
            opts
        )
        .unwrap(),
        Value::String("😀é x😀".into())
    );
    for text in [
        Value::Null,
        Value::Undefined,
        Value::Bool(false),
        Value::I32(42),
    ] {
        for chars in [Value::Null, Value::Bool(false), Value::I32(4)] {
            assert_eq!(
                run(true, text.clone(), chars.clone(), true, opts),
                run(false, text.clone(), chars, true, opts)
            );
        }
        assert_eq!(
            run(true, text.clone(), Value::Null, false, opts),
            run(false, text, Value::Null, false, opts)
        );
    }
    for old in [true, false] {
        for explicit in [true, false] {
            for defaults in [Some(DEFAULT), Some("\u{fffd}")] {
                wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                    .validate_all(&vybe_platform_wasm::write_wasm(&program(
                        old,
                        explicit,
                        strings::TrimOptions::both(defaults),
                        3,
                    )))
                    .unwrap();
            }
        }
    }
}
#[test]
fn static_ascii_defaults_eliminate_allocated_character_and_search_calls() {
    use std::sync::{
        atomic::{AtomicUsize, Ordering},
        Arc,
    };
    for old in [true, false] {
        for explicit in [true, false] {
            let mut runtime = vm(Value::String("  x  ".into()), Value::String(" ".into()));
            let chars = Arc::new(AtomicUsize::new(0));
            let searches = Arc::new(AtomicUsize::new(0));
            let observed = chars.clone();
            runtime.register_host_fn(
                "ecma:string",
                "charAt",
                Box::new(move |_ctx, args| {
                    observed.fetch_add(1, Ordering::Relaxed);
                    let unit = args[0]
                        .as_str()
                        .encode_utf16()
                        .nth(args[1].as_i32() as usize)
                        .unwrap();
                    Value::String(String::from_utf16_lossy(&[unit]).into())
                }),
            );
            let observed = searches.clone();
            runtime.register_host_fn(
                "ecma:string",
                "indexOf",
                Box::new(move |_ctx, args| {
                    observed.fetch_add(1, Ordering::Relaxed);
                    Value::F64(
                        args[0]
                            .as_str()
                            .find(args[1].as_str())
                            .map(|n| n as f64)
                            .unwrap_or(-1.0),
                    )
                }),
            );
            assert_eq!(
                runtime
                    .run(program(
                        old,
                        explicit,
                        strings::TrimOptions::both(Some(" ")),
                        1
                    ))
                    .unwrap(),
                Value::String("x".into())
            );
            let expected = if old || explicit { 6 } else { 0 };
            assert_eq!(chars.load(Ordering::Relaxed), expected);
            assert_eq!(searches.load(Ordering::Relaxed), expected);
        }
    }
}
#[test]
#[ignore = "interleaved opt-level-0 diagnostic benchmark"]
fn benchmark_static_ascii_trim_defaults() {
    use std::time::Instant;
    for n in [16, 256, 2048] {
        let text = format!("{}content{}", " \t\n".repeat(n), "\r\0\u{b}".repeat(n));
        let before = program(true, false, strings::TrimOptions::both(Some(DEFAULT)), 1);
        let after = program(false, false, strings::TrimOptions::both(Some(DEFAULT)), 1);
        let mut old = vec![];
        let mut new = vec![];
        for sample in 0..5 {
            for previous in if sample % 2 == 0 {
                [true, false]
            } else {
                [false, true]
            } {
                let mut runtime = vm(Value::String(text.clone().into()), Value::Null);
                let chunks = if previous {
                    before.clone()
                } else {
                    after.clone()
                };
                let start = Instant::now();
                assert_eq!(
                    runtime.run(chunks).unwrap(),
                    Value::String("content".into())
                );
                (if previous { &mut old } else { &mut new }).push(start.elapsed().as_micros());
            }
        }
        old.sort_unstable();
        new.sort_unstable();
        println!(
            "ASCII trim trimmed_units={} samples=5 old_us={} new_us={} speedup={:.2}x",
            n * 6,
            old[2],
            new[2],
            old[2] as f64 / new[2] as f64
        );
    }
}
