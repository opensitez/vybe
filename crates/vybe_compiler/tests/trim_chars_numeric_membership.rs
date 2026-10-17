use vybe_compiler::primitives::{self, strings};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/trim_chars_dynamic_membership.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool, left: bool, right: bool, default: Option<&'static str>) -> Vec<Chunk> {
    let mut root = Chunk::new("trim-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("trim");
    c.arity = if default.is_some() { 1 } else { 2 };
    c.param_count = c.arity;
    c.local_count = c.arity as u16;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    if default.is_none() {
        c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    }
    let argc = c.arity;
    let opts = strings::TrimOptions {
        left,
        right,
        default_chars: default,
    };
    let mut cs = vec![root, c];
    if fast {
        strings::emit_trim_chars(&mut cs, 1, argc, opts, 0);
    } else {
        before::emit_trim_chars(&mut cs, 1, argc, opts, 0);
    }
    cs[1].emit_op(Op::RETURN, 0);
    primitives::bundle::finalize_with_runtime_helpers(&mut cs);
    primitives::globals::declare_free_globals(&mut cs);
    primitives::globals::normalize_global_table(&mut cs);
    let bindings: Vec<_> = cs[0]
        .global_inits
        .iter()
        .filter_map(|init| {
            if let vybe_runtime::chunk::ConstExpr::RefFunc(index) = init.init {
                Some((init.name.clone(), index))
            } else {
                None
            }
        })
        .collect();
    for (name, index) in bindings {
        let mut f = Chunk::new(format!("binding-factory:{name}"));
        f.emit_op_u16(Op::REF_FUNC, index as u16, 0);
        f.emit(0, 0);
        f.emit_op(Op::RETURN, 0);
        cs.push(f);
    }
    cs
}
fn worker(fast: bool, left: bool, right: bool, default: Option<&'static str>) -> (VM, Value) {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(program(fast, left, right, default)).unwrap();
    (vm, f)
}
#[test]
fn trim_preserves_explicit_and_nonascii_sets() {
    let mut checks = 0;
    for (left, right) in [(true, true), (true, false), (false, true), (false, false)] {
        for default in [None, Some("é😀"), Some(" \t\n\r\0\u{b}")] {
            let mut workers = [
                worker(false, left, right, default),
                worker(true, left, right, default),
            ];
            for input in [
                "",
                "a",
                "aaaa",
                " aaabaaa ",
                "\0x\0",
                "éabcé",
                "😀x😀",
                "é😀尾😀é",
                "\u{fffd}x\u{fffd}",
                "[]$.*x[]$.*",
            ] {
                for set in ["", "a", " a", "[]$.*", "\0", "é😀", "\u{fffd}"] {
                    let args = if default.is_some() {
                        vec![Value::String(input.into())]
                    } else {
                        vec![Value::String(input.into()), Value::String(set.into())]
                    };
                    let result = workers
                        .each_mut()
                        .map(|(vm, f)| vm.invoke(f, &args).map_err(|e| e.message));
                    assert_eq!(
                        result[0], result[1],
                        "{input:?}, {set:?}, {left}/{right}, {default:?}"
                    );
                    // Independent oracle where Rust scalar characters match the UTF-16 membership.
                    let set = default.unwrap_or(set);
                    if !input.contains('😀') && !set.contains('😀') {
                        let mut expected = input;
                        if left {
                            expected = expected.trim_start_matches(|c| set.contains(c));
                        }
                        if right {
                            expected = expected.trim_end_matches(|c| set.contains(c));
                        }
                        assert_eq!(result[1], Ok(Value::String(expected.into())));
                    }
                    checks += 1;
                }
            }
        }
        for value in [
            Value::Null,
            Value::Undefined,
            Value::Bool(true),
            Value::I32(17),
            Value::bigint_i64(2),
        ] {
            for args in [
                [Value::String(" abc ".into()), value.clone()],
                [value, Value::String(" ".into())],
            ] {
                let mut outputs = vec![];
                for fast in [false, true] {
                    let (mut vm, f) = worker(fast, left, right, None);
                    outputs.push(vm.invoke(&f, &args).map_err(|e| e.message));
                }
                assert_eq!(outputs[0], outputs[1]);
                checks += 1;
            }
        }
    }
    eprintln!("Verified {checks} trim cases");
}
#[test]
fn size_and_native_exports() {
    let mut rows = vec![];
    for (label, left, right, default) in [
        ("both", true, true, None),
        ("start", true, false, None),
        ("end", false, true, None),
        ("neither", false, false, None),
        ("unicode-default", true, true, Some("é😀")),
        ("ascii-default", true, true, Some(" \t\n\r\0\u{b}")),
    ] {
        let mut sizes = vec![];
        for fast in [false, true] {
            let cs = program(fast, left, right, default);
            let bytes = vybe_platform_wasm::write_wasm_with_options(
                &cs,
                vybe_platform_wasm::WasmWriteOptions {
                    include_name_section: false,
                    include_vybe_metadata: false,
                },
            );
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&bytes)
                .unwrap();
            if std::env::var_os("VYBE_PRIMITIVE_TRIM_EXPORT").is_some() {
                native_export::export_chunks(&cs, &format!("trim-{label}-{}", u8::from(fast)));
            }
            sizes.push(serde_json::json!({"body":cs[1].code.len(),"native":bytes.len(),"locals":cs[1].local_count}));
        }
        if label == "neither" || label == "ascii-default" {
            assert_eq!(sizes[0], sizes[1]);
        } else {
            assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
            assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
        }
        rows.push(serde_json::json!({"case":label,"sizes":sizes}));
    }
    if std::env::var_os("VYBE_PRIMITIVE_TRIM_EXPORT").is_some() {
        std::fs::write(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../TRIM_GATES.json"),
            serde_json::to_vec_pretty(&rows).unwrap(),
        )
        .unwrap();
    }
    eprintln!("{rows:?}");
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::{Duration, Instant};
    let mut rows = vec![];
    for (label, text, set) in [
        ("empty", String::new(), "a"),
        ("unchanged-8", "abcdefgh".into(), "x"),
        ("trim-8", "aaaxaaaa".into(), "a"),
        (
            "trim-4096",
            format!("{}x{}", "a".repeat(2048), "a".repeat(2048)),
            "a",
        ),
        (
            "unicode-2048",
            format!("{}x{}", "é".repeat(1024), "é".repeat(1024)),
            "é",
        ),
    ] {
        let mut workers = [
            worker(false, true, true, None),
            worker(true, true, true, None),
        ];
        let args = [Value::String(text.into()), Value::String(set.into())];
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let (vm, f) = &mut workers[side];
                let start = Instant::now();
                let mut n = 0;
                while start.elapsed() < Duration::from_millis(100) || n < 8 {
                    std::hint::black_box(vm.invoke(f, &args).unwrap());
                    n += 1;
                }
                samples[side].push(start.elapsed().as_secs_f64() / n as f64);
            }
        }
        let mut ratios: Vec<_> = samples[0]
            .iter()
            .zip(&samples[1])
            .map(|(a, b)| a / b)
            .collect();
        ratios.sort_by(f64::total_cmp);
        rows.push(
            serde_json::json!({"case":label,"median_speedup":ratios[3],"seconds_per_call":samples}),
        );
    }
    let result = serde_json::json!({"profile":"opt0 assertions and overflow checks","pairs":7,"minimum_sample_ms":100,"minimum_calls_per_sample":8,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../TRIM_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
