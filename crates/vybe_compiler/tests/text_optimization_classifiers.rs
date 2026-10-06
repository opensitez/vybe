use vybe_compiler::primitives::{bundle, globals, platforms, strings};
use vybe_runtime::{Chunk, Op, Value, VM};

#[path = "fixtures/text_optimization_classifiers.rs"]
mod baseline;

type Emitter = fn(&mut [Chunk], usize, u32);
const CURRENT: [Emitter; 6] = [
    strings::emit_is_digit,
    strings::emit_is_alpha,
    strings::emit_is_alnum,
    strings::emit_is_space,
    strings::emit_is_upper,
    strings::emit_is_lower,
];
const BASELINE: [Emitter; 6] = [
    baseline::emit_is_digit,
    baseline::emit_is_alpha,
    baseline::emit_is_alnum,
    baseline::emit_is_space,
    baseline::emit_is_upper,
    baseline::emit_is_lower,
];

fn program(old: bool, class: usize, calls: usize) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<text-classifiers>")];
    for call in 0..calls {
        globals::emit_read(&mut chunks[0], "text", 0);
        (if old { BASELINE } else { CURRENT })[class](&mut chunks, 0, 0);
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
fn vm(text: &str) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("text", Value::String(text.into()));
    vm
}

fn run_cached(runtime: &mut VM, chunks: &[Chunk], text: &str) -> Value {
    runtime.set_global("text", Value::String(text.into()));
    // Each run resets execution state and relocates a fresh program copy.
    runtime.run(chunks.to_vec()).unwrap()
}

#[test]
fn ascii_ranges_preserve_every_ascii_character_and_unicode_fallback() {
    for class in 0..6 {
        let old = program(true, class, 1);
        let new = program(false, class, 1);
        let mut old_vm = vm("");
        let mut new_vm = vm("");
        for cp in 0..128 {
            let text = char::from_u32(cp).unwrap().to_string();
            assert_eq!(
                run_cached(&mut old_vm, &old, &text),
                run_cached(&mut new_vm, &new, &text),
                "class={class} cp={cp}"
            );
        }
        for text in [
            "", "ABC", "abc", "ABC!", "aA", "123", "1é", "é", "É", "éÉ", "世", "😀", "𝔸", "\u{a0}",
            "\u{2003}", "\u{fffd}", "e\u{301}", "\0", " \t\n", "\u{0661}", "\u{2160}",
        ] {
            assert_eq!(
                run_cached(&mut old_vm, &old, text),
                run_cached(&mut new_vm, &new, text),
                "class={class} text={text:?}"
            );
        }
        for cp in [
            127, 128, 255, 256, 0xd7ff, 0xe000, 0xffff, 0x10000, 0x10ffff,
        ] {
            let text = char::from_u32(cp).unwrap().to_string();
            assert_eq!(
                run_cached(&mut old_vm, &old, &text),
                run_cached(&mut new_vm, &new, &text),
                "boundary class={class} cp={cp}"
            );
        }
        for old in [true, false] {
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&vybe_platform_wasm::write_wasm(&program(old, class, 3)))
                .unwrap();
        }
    }
}

#[test]
fn ascii_classification_skips_unicode_table_reads() {
    use std::sync::{
        atomic::{AtomicUsize, Ordering},
        Arc,
    };
    for (class, text) in [
        (0, "0123456789"),
        (1, "Alphabetic"),
        (2, "Abc123"),
        (3, " \t\n\r"),
        (4, "ABC 123"),
        (5, "abc 123"),
    ] {
        for old in [true, false] {
            let count = Arc::new(AtomicUsize::new(0));
            let observed = count.clone();
            let mut runtime = vm(text);
            runtime.register_host_fn(
                "wasm:js-string",
                "charCodeAt",
                Box::new(move |_ctx, args| {
                    observed.fetch_add(1, Ordering::Relaxed);
                    let unit = args[0]
                        .as_str()
                        .encode_utf16()
                        .nth(args[1].as_i32() as usize)
                        .unwrap();
                    Value::I32(unit as i32)
                }),
            );
            runtime.run(program(old, class, 1)).unwrap();
            if old {
                assert!(count.load(Ordering::Relaxed) > 0);
            } else {
                assert_eq!(count.load(Ordering::Relaxed), 0);
            }
        }
    }
}

#[test]
#[ignore = "interleaved opt-level-0 diagnostic benchmark"]
fn benchmark_ascii_classification() {
    use std::time::Instant;
    for (class, label, pattern) in [
        (0, "digit", "0123456789"),
        (1, "alpha", "Alphabetic"),
        (4, "upper", "UPPER CASE 123"),
    ] {
        for units in [256, 2048] {
            let text = pattern.repeat(units / pattern.len() + 1)[..units].to_owned();
            let old = program(true, class, 1);
            let new = program(false, class, 1);
            let mut before = Vec::new();
            let mut after = Vec::new();
            for sample in 0..5 {
                for previous in if sample % 2 == 0 {
                    [true, false]
                } else {
                    [false, true]
                } {
                    let mut runtime = vm(&text);
                    let chunks = if previous { old.clone() } else { new.clone() };
                    let start = Instant::now();
                    runtime.run(chunks).unwrap();
                    (if previous { &mut before } else { &mut after })
                        .push(start.elapsed().as_micros());
                }
            }
            before.sort_unstable();
            after.sort_unstable();
            println!(
                "class={label} units={units} samples=5 old_us={} new_us={} speedup={:.2}x",
                before[2],
                after[2],
                before[2] as f64 / after[2] as f64
            );
        }
    }
}
