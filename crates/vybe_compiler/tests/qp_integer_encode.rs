//! Deferred integer encoder candidate: production retains the accepted baseline.
use vybe_compiler::primitives;
#[path = "fixtures/qp_encode_integer_candidate.rs"]
mod candidate;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/qp_encode_integer_before.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("qp-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("entry");
    c.arity = 1;
    c.local_count = 1;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let mut bodies = vec![c];
    if fast {
        candidate::emit_quoted_printable_encode(&mut bodies, 0, 1, 0);
    } else {
        before::emit_quoted_printable_encode(&mut bodies, 0, 1, 0);
    }
    bodies[0].emit_op(Op::RETURN, 0);
    let mut c = bodies.remove(0);
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
    let hex = b"0123456789ABCDEF";
    let mut out = String::new();
    for u in s.encode_utf16() {
        if u == 61 || u > 126 || (u < 32 && u != 9) {
            out.push('=');
            if let Some(&h) = hex.get((u / 16) as usize) {
                out.push(h as char);
            }
            out.push(hex[(u % 16) as usize] as char);
        } else {
            out.push(char::from_u32(u as u32).unwrap());
        }
    }
    out
}
fn check(ws: &mut [(VM, Value); 2], s: &str) -> usize {
    let expected = Value::String(oracle(s).into());
    for (vm, f) in ws {
        assert_eq!(
            vm.invoke(f, &[Value::String(s.into())]).unwrap(),
            expected,
            "input={s:?}"
        );
    }
    2
}
#[test]
fn exhaustive_units_and_sizes() {
    let mut ws = [worker(false), worker(true)];
    let mut checks = 0;
    let mut domain = 0;
    for base in (0..65536u32).step_by(256) {
        let s: String = (base..base + 256)
            .filter_map(char::from_u32)
            .inspect(|_| domain += 1)
            .collect();
        checks += check(&mut ws, &s);
    }
    let alphabet = ['=', 'A', '\r', '\n', ' ', '\t', 'é', '中'];
    for len in 0..5 {
        for mut bits in 0..8usize.pow(len) {
            let s: String = (0..len)
                .map(|_| {
                    let ch = alphabet[bits % 8];
                    bits /= 8;
                    ch
                })
                .collect();
            checks += check(&mut ws, &s);
        }
    }
    for s in [
        "",
        "=",
        "\0\t\n\r ",
        "Abc=éÿ中😀",
        "\u{10000}\u{10ffff}",
        "\u{7f}\u{80}\u{ff}\u{100}",
    ] {
        checks += check(&mut ws, s);
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
        native_export::export_chunks(&cs, &format!("qp-encode-integer-{fast}"));
        sizes.push(serde_json::json!({"body":cs[1].code.len(),"locals":cs[1].local_count,"native":bytes.len()}));
    }
    assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
    assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
    assert!(sizes[1]["locals"].as_u64() < sizes[0]["locals"].as_u64());
    let result = serde_json::json!({"checks":checks,"scalar_domain":domain,"sizes":sizes});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../QP_ENCODE_INTEGER_GATES.json"),
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
        ("control", "abcd".repeat(64), true),
        ("empty", String::new(), false),
        ("short", "=\r\n".into(), false),
        ("plain256", "abcd".repeat(64), false),
        ("escapes256", "=\r\n\0".repeat(64), false),
        ("latin256", "éÿ".repeat(128), false),
        ("unicode256", "中😀\u{100}".repeat(64), false),
        ("mixed", "Abc=é中😀\0\t\r\n".repeat(128), false),
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
            .join("../../QP_ENCODE_INTEGER_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}

#[test]
fn physical_dispatches() {
    let mut rows = vec![];
    for (label, text) in [
        ("empty", String::new()),
        ("plain256", "abcd".repeat(64)),
        ("escapes256", "=\r\n\0".repeat(64)),
        ("mixed", "Abc=é中😀\0\t\r\n".repeat(128)),
    ] {
        let mut dispatches = vec![];
        for fast in [false, true] {
            let (mut vm, f) = worker(fast);
            let input = Value::String(text.clone().into());
            for _ in 0..8 {
                vm.invoke(&f, &[input.clone()]).unwrap();
            }
            vm.record_runtime_perf(true);
            assert_eq!(
                vm.invoke(&f, &[input]).unwrap(),
                Value::String(oracle(&text).into())
            );
            dispatches.push(vm.take_runtime_perf().unwrap().dispatches);
        }
        rows.push(serde_json::json!({"label":label,"units":text.encode_utf16().count(),"dispatches":dispatches}));
    }
    let result = serde_json::json!({"scope":"Warm exact helper invocation; physical runtime dispatch count is separate from wall-clock timings.","rows":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../QP_ENCODE_INTEGER_DISPATCHES.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
