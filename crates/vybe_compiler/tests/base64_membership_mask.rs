#![cfg(feature = "primitive-benchmarks")]
use vybe_compiler::primitives;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/base64_mask_before.rs"]
mod before;
#[path = "fixtures/base64_mask_candidate.rs"]
mod candidate;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool, url: bool, repeats: usize) -> Vec<Chunk> {
    program_with(
        if fast {
            primitives::base64::emit_code_is_allowed
        } else {
            before::emit_code_is_allowed
        },
        url,
        repeats,
    )
}
fn program_with(emit: fn(&mut Chunk, u16, bool, u32), url: bool, repeats: usize) -> Vec<Chunk> {
    let mut root = Chunk::new("base64-membership-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("entry");
    c.arity = 1;
    c.param_count = 1;
    c.local_count = 1;
    for i in 0..repeats {
        emit(&mut c, 0, url, 0);
        if i + 1 < repeats {
            c.emit_op(Op::DROP, 0);
        }
    }
    c.emit_op(Op::RETURN, 0);
    let mut cs = vec![root, c];
    primitives::globals::declare_free_globals(&mut cs);
    primitives::globals::normalize_global_table(&mut cs);
    cs
}
fn vm() -> VM {
    let mut v = VM::new();
    primitives::platforms::register_platforms_all(&mut v);
    v
}
fn output(name: &str, value: &serde_json::Value) {
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join(format!("../../BASE64_MEMBERSHIP_{name}.json")),
        serde_json::to_vec_pretty(value).unwrap(),
    )
    .unwrap();
}
fn oracle(code: i32, url: bool) -> i32 {
    i32::from(
        matches!(code,48..=57|65..=90|97..=122|61)
            || if url {
                code == 45 || code == 95
            } else {
                code == 43 || code == 47
            },
    )
}
#[test]
fn exhaustive_units_and_size() {
    let mut checks = 0;
    let mut rows = vec![];
    for url in [false, true] {
        let production = program(true, url, 1);
        let measured = program_with(candidate::emit_code_is_allowed, url, 1);
        assert_eq!(production[1].code, measured[1].code);
        assert_eq!(production[1].imports, measured[1].imports);
        assert_eq!(production[1].constants, measured[1].constants);
        assert_eq!(production[1].local_count, measured[1].local_count);
        let mut v = vm();
        let fs = [
            v.run(program(false, url, 1)).unwrap(),
            v.run(program(true, url, 1)).unwrap(),
        ];
        for code in (0..65536).chain([i32::MIN, -1, -32, 65536, 0x110000, i32::MAX]) {
            for f in &fs {
                assert_eq!(
                    v.invoke(f, &[Value::I32(code)]).unwrap(),
                    Value::I32(oracle(code, url)),
                    "url={url} code={code}"
                );
                checks += 1;
            }
        }
        let mut sizes = vec![];
        for fast in [false, true] {
            let cs = program(fast, url, 1);
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
            native_export::export_chunks(&cs, &format!("base64-membership-{url}-{fast}"));
            // Preserve the primitive's import-free numeric-body contract too.
            let mut raw = cs[1].clone();
            raw.func_sig = Some(("i32".into(), "i32".into()));
            assert!(raw.imports.is_empty());
            let raw_bytes = vybe_platform_wasm::write_wasm(&[raw.clone()]);
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&raw_bytes)
                .unwrap();
            native_export::export_chunks(&[raw], &format!("base64-membership-raw-{url}-{fast}"));
            sizes.push(serde_json::json!({"body":cs[1].code.len(),"locals":cs[1].local_count,"native":bytes.len(),"raw_native":raw_bytes.len()}));
        }
        assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
        assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
        assert!(sizes[1]["raw_native"].as_u64() < sizes[0]["raw_native"].as_u64());
        let mut dispatches = vec![];
        for code in [65, 48, 61, 45, 47, 33, 128, 65535] {
            let mut counts = vec![];
            for f in &fs {
                v.record_runtime_perf(true);
                v.invoke(f, &[Value::I32(code)]).unwrap();
                counts.push(v.take_runtime_perf().unwrap().dispatches);
            }
            dispatches.push(serde_json::json!({"code":code,"counts":counts}));
        }
        rows.push(serde_json::json!({"url":url,"sizes":sizes,"dispatches":dispatches}));
    }
    let result = serde_json::json!({"checks":checks,"rows":rows});
    output("GATES", &result);
    eprintln!("{result}");
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for (label, code, url, control) in [
        ("control", 65, false, true),
        ("letter", 65, false, false),
        ("digit", 48, false, false),
        ("padding", 61, false, false),
        ("punctuation", 33, false, false),
        ("nonascii", 128, false, false),
        ("url-letter", 65, true, false),
        ("url-underscore", 95, true, false),
    ] {
        let mut v = vm();
        let a = v.run(program(false, url, 32)).unwrap();
        let b = if control {
            a.clone()
        } else {
            v.run(program(true, url, 32)).unwrap()
        };
        let fs = [a, b];
        let args = [Value::I32(code)];
        for f in &fs {
            for _ in 0..8 {
                std::hint::black_box(v.invoke(f, &args).unwrap());
            }
        }
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            let mut seconds = [0.0, 0.0];
            let mut n = 0;
            let mut turn = pair;
            while seconds[0] < 0.5 || seconds[1] < 0.5 || n < 8 {
                for side in if turn % 2 == 0 { [0, 1] } else { [1, 0] } {
                    let t = Instant::now();
                    std::hint::black_box(v.invoke(&fs[side], &args).unwrap());
                    seconds[side] += t.elapsed().as_secs_f64();
                }
                n += 1;
                turn += 1;
            }
            for side in 0..2 {
                samples[side].push(seconds[side] / n as f64);
            }
        }
        let mut ratios: Vec<_> = samples[0]
            .iter()
            .zip(&samples[1])
            .map(|(a, b)| a / b)
            .collect();
        ratios.sort_by(f64::total_cmp);
        rows.push(serde_json::json!({"label":label,"url":url,"control":control,"code":code,"repeats":32,"median_speedup":ratios[3],"seconds_per_call":samples}));
    }
    let r = serde_json::json!({"same_vm":true,"equal_calls_interleaved":true,"pairs":7,"minimum_sample_ms":500,"cases":rows});
    output("MEASUREMENTS", &r);
    eprintln!("{r}");
}
