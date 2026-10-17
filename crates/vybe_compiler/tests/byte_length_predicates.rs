#![cfg(feature = "primitive-benchmarks")]
use vybe_compiler::primitives;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/byte_length_before.rs"]
mod before;
#[path = "fixtures/byte_length_candidate.rs"]
mod candidate;
#[path = "fixtures/byte_length_bulk_before.rs"]
mod historical;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool) -> Vec<Chunk> {
    program_with(if fast {
        historical::emit_byte_length
    } else {
        before::emit_byte_length
    })
}
fn program_with(emit: fn(&mut [Chunk], usize, u32)) -> Vec<Chunk> {
    let mut root = Chunk::new("byte-length-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("entry");
    c.arity = 1;
    c.param_count = 1;
    c.local_count = 1;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let mut cs = vec![root, c];
    emit(&mut cs, 1, 0);
    cs[1].emit_op(Op::RETURN, 0);
    primitives::globals::declare_free_globals(&mut cs);
    primitives::globals::normalize_global_table(&mut cs);
    cs
}
fn vm() -> VM {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    vm
}
fn output(name: &str, value: &serde_json::Value) {
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join(format!("../../BYTE_LENGTH_{name}.json")),
        serde_json::to_vec_pretty(value).unwrap(),
    )
    .unwrap();
}
#[test]
fn unicode_values_sizes_and_dispatches() {
    // Preserve this historical predicate campaign against its frozen recipe.
    // Current production and stringref are covered by byte_length_bulk.
    let measured = program_with(candidate::emit_byte_length);
    let production = program(true);
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
    let mut check = |s: String| {
        for f in &fs {
            assert_eq!(
                v.invoke(f, &[Value::String(s.clone().into())]).unwrap(),
                Value::I32(s.len() as i32)
            );
            checks += 1;
        }
    };
    for base in (0..0x110000u32).step_by(1024) {
        let s: String = (base..base + 1024).filter_map(char::from_u32).collect();
        check(s);
    }
    for cp in [
        0, 1, 126, 127, 128, 129, 2047, 2048, 2049, 0xd7ff, 0xe000, 65535, 65536, 65537, 0x10ffff,
    ] {
        check(char::from_u32(cp).unwrap().to_string());
    }
    for s in ["", "hello", "a\0b", "é界😀é", "😀😀", "a😀b界"] {
        check(s.into());
    }
    let mut sizes = vec![];
    let mut dispatches = vec![];
    for (fast, f) in [false, true].into_iter().zip(&fs) {
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
        native_export::export_chunks(&cs, &format!("byte-length-{fast}"));
        sizes.push(serde_json::json!({"body":cs[1].code.len(),"locals":cs[1].local_count,"native":bytes.len()}));
        v.record_runtime_perf(true);
        v.invoke(f, &[Value::String("aé界😀".repeat(16).into())])
            .unwrap();
        dispatches.push(v.take_runtime_perf().unwrap().dispatches);
    }
    assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
    assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
    assert!(dispatches[1] < dispatches[0]);
    let mut dispatch_cases = vec![];
    for text in ["", "abc", "ééé", "界界界", "😀😀😀", "aé界😀"] {
        let mut counts = vec![];
        for f in &fs {
            v.record_runtime_perf(true);
            v.invoke(f, &[Value::String(text.into())]).unwrap();
            counts.push(v.take_runtime_perf().unwrap().dispatches);
        }
        assert!(counts[1] <= counts[0], "{text:?}: {counts:?}");
        dispatch_cases.push(serde_json::json!({"text":text,"dispatches":counts}));
    }
    let result = serde_json::json!({"checks":checks,"sizes":sizes,"dispatches":dispatches,"dispatch_cases":dispatch_cases});
    output("GATES", &result);
    eprintln!("{result}");
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for (label, text, control) in [
        ("control", "abcd".repeat(64), true),
        ("empty", String::new(), false),
        ("short", "abc".into(), false),
        ("ascii", "abcd".repeat(64), false),
        ("two-byte", "é".repeat(256), false),
        ("three-byte", "界".repeat(256), false),
        ("astral", "😀".repeat(128), false),
        ("mixed", "aé界😀".repeat(64), false),
    ] {
        let mut v = vm();
        let a = v.run(program(false)).unwrap();
        let b = if control {
            a.clone()
        } else {
            v.run(program(true)).unwrap()
        };
        let fs = [a, b];
        let args = [Value::String(text.clone().into())];
        for f in &fs {
            for _ in 0..8 {
                std::hint::black_box(v.invoke(f, &args).unwrap());
            }
        }
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            let mut elapsed = [0.0, 0.0];
            let mut calls = 0;
            let mut turn = pair;
            while elapsed[0] < 0.5 || elapsed[1] < 0.5 || calls < 8 {
                for side in if turn % 2 == 0 { [0, 1] } else { [1, 0] } {
                    let start = Instant::now();
                    std::hint::black_box(v.invoke(&fs[side], &args).unwrap());
                    elapsed[side] += start.elapsed().as_secs_f64();
                }
                calls += 1;
                turn += 1;
            }
            for side in 0..2 {
                samples[side].push(elapsed[side] / calls as f64);
            }
        }
        let mut ratios: Vec<_> = samples[0]
            .iter()
            .zip(&samples[1])
            .map(|(a, b)| a / b)
            .collect();
        ratios.sort_by(f64::total_cmp);
        rows.push(serde_json::json!({"label":label,"control":control,"units":text.encode_utf16().count(),"median_speedup":ratios[3],"seconds_per_call":samples}));
    }
    let result = serde_json::json!({"same_vm":true,"equal_calls_interleaved":true,"pairs":7,"minimum_sample_ms":500,"cases":rows});
    output("MEASUREMENTS", &result);
    eprintln!("{result}");
}
