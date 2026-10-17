use vybe_compiler::primitives::{self, polyfills};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/rot13_integer_before.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("rot13-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = if fast {
        polyfills::build_runtime_helpers_for_exports(&mut root, &["__stdlib_encoding_str_rot13"])
            .chunks
            .remove(0)
    } else {
        let mut c = Chunk::new("entry");
        c.arity = 1;
        c.local_count = 1;
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        let mut cs = vec![c];
        before::emit_str_rot13(&mut cs, 0, 1, 0);
        cs[0].emit_op(Op::RETURN, 0);
        cs.remove(0)
    };
    c.name = "entry".into();
    c.param_count = c.arity;
    let mut cs = vec![root, c];
    primitives::bundle::finalize_with_runtime_helpers(&mut cs);
    primitives::globals::declare_free_globals(&mut cs);
    primitives::globals::normalize_global_table(&mut cs);
    cs
}
fn worker(fast: bool) -> (VM, Value) {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(program(fast)).unwrap();
    (vm, f)
}
fn oracle(s: &str) -> String {
    s.encode_utf16()
        .map(|u| {
            let c = if (65..=90).contains(&u) || (97..=122).contains(&u) {
                if (u | 32) <= 109 { u + 13 } else { u - 13 }
            } else {
                u
            };
            String::from_utf16_lossy(&[c])
        })
        .collect()
}
#[test]
fn exhaustive_units_and_sizes() {
    let mut ws = [worker(false), worker(true)];
    let mut checks = 0;
    for unit in 0..=65535u32 {
        if let Some(ch) = char::from_u32(unit) {
            let s = ch.to_string();
            let expected = oracle(&s);
            for (vm, f) in &mut ws {
                assert_eq!(
                    vm.invoke(f, &[Value::String(s.clone().into())]).unwrap(),
                    Value::String(expected.clone().into()),
                    "unit={unit}"
                );
                checks += 1;
            }
        }
    }
    for s in [
        "",
        "Abz 13! ÿ",
        "A\0Nz mM",
        "é😀Abc😀é",
        "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz",
        "{[\\^_`]}",
    ] {
        for (vm, f) in &mut ws {
            assert_eq!(
                vm.invoke(f, &[Value::String(s.into())]).unwrap(),
                Value::String(oracle(s).into())
            );
            checks += 1;
        }
    }
    for input in [
        Value::Null,
        Value::Undefined,
        Value::I32(1),
        Value::Bool(true),
    ] {
        let outcomes = [false, true].map(|fast| {
            let (mut vm, f) = worker(fast);
            vm.invoke(&f, &[input.clone()])
                .unwrap_err()
                .to_string()
                .contains("TypeError")
        });
        assert_eq!(outcomes, [true, true]);
        checks += 2;
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
        native_export::export_chunks(&cs, &format!("rot13-integer-{fast}"));
        sizes.push(serde_json::json!({"body":cs[1].code.len(),"locals":cs[1].local_count,"native":bytes.len()}));
    }
    assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
    assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
    assert!(sizes[1]["locals"].as_u64() < sizes[0]["locals"].as_u64());
    let result = serde_json::json!({"checks":checks,"sizes":sizes});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../ROT13_INTEGER_GATES.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::{Duration, Instant};
    let mut rows = vec![];
    for (label, text, control) in [
        ("control", "Abz ".repeat(64), true),
        ("empty", String::new(), false),
        ("short", "Abz 13!".into(), false),
        ("letters256", "AbzNopqRstUvwXyZ".repeat(16), false),
        ("nonletters256", "é012!_[]".repeat(32), false),
        ("mixed4096", "Abz 13! é中😀".repeat(320), false),
    ] {
        let (mut vm, a) = worker(false);
        let b = if control {
            a.clone()
        } else {
            vm.run(program(true)).unwrap()
        };
        let fs = [a, b];
        let input = Value::String(text.clone().into());
        for f in &fs {
            for _ in 0..8 {
                std::hint::black_box(vm.invoke(f, &[input.clone()]).unwrap());
            }
        }
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let t = Instant::now();
                let mut n = 0;
                while t.elapsed() < Duration::from_millis(500) || n < 8 {
                    std::hint::black_box(vm.invoke(&fs[side], &[input.clone()]).unwrap());
                    n += 1;
                }
                samples[side].push(t.elapsed().as_secs_f64() / n as f64);
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
    let result = serde_json::json!({"same_vm":true,"pairs":7,"minimum_sample_ms":500,"minimum_calls":8,"warmups":8,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../ROT13_INTEGER_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
