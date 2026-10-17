#![cfg(feature = "primitive-benchmarks")]
use vybe_compiler::primitives;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/base64_mask_before.rs"]
mod before;
#[path = "fixtures/base64_mask_candidate.rs"]
mod candidate;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
#[path = "fixtures/base64_strict_before.rs"]
mod strict_before;
#[path = "fixtures/base64_strict_candidate.rs"]
mod strict_candidate;
fn program(fast: bool) -> Vec<Chunk> {
    if !fast {
        return program_fixture(false);
    }
    let mut root = Chunk::new("base64-strict-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = primitives::polyfills::build_runtime_helpers_for_exports(
        &mut root,
        &["__stdlib_encoding_base64_decode"],
    )
    .chunks
    .remove(0);
    c.name = "entry".into();
    c.param_count = c.arity;
    let mut cs = vec![root, c];
    primitives::bundle::finalize_with_runtime_helpers(&mut cs);
    primitives::globals::declare_free_globals(&mut cs);
    primitives::globals::normalize_global_table(&mut cs);
    cs
}
fn program_fixture(fast: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("base64-strict-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("entry");
    c.arity = 2;
    c.param_count = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let mut cs = vec![root, c];
    if fast {
        strict_candidate::emit_base64_decode(&mut cs, 1, 2, 0)
    } else {
        strict_before::emit_base64_decode(&mut cs, 1, 2, 0)
    }
    cs[1].emit_op(Op::RETURN, 0);
    primitives::bundle::finalize_with_runtime_helpers(&mut cs);
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
            .join(format!("../../BASE64_STRICT_{name}.json")),
        serde_json::to_vec_pretty(value).unwrap(),
    )
    .unwrap();
}
#[test]
fn values_sizes_work() {
    let production = program(true);
    let measured = program_fixture(true);
    assert_eq!(production[1].code, measured[1].code);
    assert_eq!(production[1].imports, measured[1].imports);
    assert_eq!(production[1].constants, measured[1].constants);
    assert_eq!(production[1].local_count, measured[1].local_count);
    let mut v = vm();
    let fs = [
        v.run(program(false)).unwrap(),
        v.run(program(true)).unwrap(),
    ];
    let mut checks = 0;
    for cp in 0..65536u32 {
        if matches!(cp,48..=57|65..=90|97..=122|43|47|61) {
            continue;
        }
        if let Some(ch) = char::from_u32(cp) {
            let args = [Value::String(format!("AA{ch}A").into()), Value::Bool(true)];
            for f in &fs {
                assert_eq!(v.invoke(f, &args).unwrap(), Value::Bool(false), "code={cp}");
                checks += 1;
            }
        }
    }
    for (s, expected) in [
        ("", ""),
        ("AA==", "\0"),
        ("TWFu", "Man"),
        ("aGVsbG8=", "hello"),
        ("YWJjYWJj", "abcabc"),
    ] {
        for f in &fs {
            assert_eq!(
                v.invoke(f, &[Value::String(s.into()), Value::Bool(true)])
                    .unwrap(),
                Value::String(expected.into())
            );
            checks += 1;
        }
    }
    for (s, expected) in [("T W\nFu", "Man"), ("aGV!sbG8=", "hello"), ("", "")] {
        for strict in [
            Value::Bool(false),
            Value::Null,
            Value::Undefined,
            Value::I32(0),
            Value::String("".into()),
        ] {
            for f in &fs {
                assert_eq!(
                    v.invoke(f, &[Value::String(s.into()), strict.clone()])
                        .unwrap(),
                    Value::String(expected.into())
                );
                checks += 1;
            }
        }
    }
    let mut sizes = vec![];
    for fast in [false, true] {
        let cs = program(fast);
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
        native_export::export_chunks(&cs, &format!("base64-strict-{fast}"));
        sizes.push(serde_json::json!({"body":cs[1].code.len(),"locals":cs[1].local_count,"native":bytes.len()}));
    }
    let mut work = vec![];
    for (label, s) in [
        ("empty", String::new()),
        ("valid", "YWJj".repeat(256)),
        ("early", format!("!{}", "A".repeat(4095))),
        (
            "middle",
            format!("{}!{}", "A".repeat(2048), "A".repeat(2047)),
        ),
        ("late", format!("{}!", "A".repeat(4095))),
    ] {
        let mut counts = vec![];
        for f in &fs {
            let args = [Value::String(s.clone().into()), Value::Bool(true)];
            v.invoke(f, &args).unwrap();
            v.record_runtime_perf(true);
            v.invoke(f, &args).unwrap();
            counts.push(v.take_runtime_perf().unwrap().dispatches);
        }
        assert!(counts[1] <= counts[0]);
        work.push(serde_json::json!({"label":label,"units":s.len(),"dispatches":counts}));
    }
    assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
    assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
    let r = serde_json::json!({"checks":checks,"sizes":sizes,"work":work});
    output("GATES", &r);
    eprintln!("{r}");
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::{Duration, Instant};
    let mut rows = vec![];
    for (label, s, control) in [
        ("control", "YWJj".repeat(64), true),
        ("empty", String::new(), false),
        ("valid-short", "TWFu".into(), false),
        ("valid-long", "YWJj".repeat(256), false),
        ("invalid-early-short", format!("!{}", "A".repeat(63)), false),
        (
            "invalid-early-long",
            format!("!{}", "A".repeat(4095)),
            false,
        ),
        (
            "invalid-middle",
            format!("{}!{}", "A".repeat(2048), "A".repeat(2047)),
            false,
        ),
        ("invalid-late", format!("{}!", "A".repeat(4095)), false),
    ] {
        let mut v = vm();
        let a = v.run(program(false)).unwrap();
        let b = if control {
            a.clone()
        } else {
            v.run(program(true)).unwrap()
        };
        let fs = [a, b];
        let args = [Value::String(s.clone().into()), Value::Bool(true)];
        for f in &fs {
            for _ in 0..8 {
                std::hint::black_box(v.invoke(f, &args).unwrap());
            }
        }
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let start = Instant::now();
                let mut n = 0;
                while start.elapsed() < Duration::from_millis(500) || n < 8 {
                    std::hint::black_box(v.invoke(&fs[side], &args).unwrap());
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
        rows.push(serde_json::json!({"label":label,"control":control,"units":s.len(),"median_speedup":ratios[3],"seconds_per_call":samples}));
    }
    let r = serde_json::json!({"same_vm":true,"pairs":7,"alternating_fixed_time_samples":true,"minimum_sample_ms":500,"minimum_calls":8,"cases":rows});
    output("MEASUREMENTS", &r);
    eprintln!("{r}");
}
