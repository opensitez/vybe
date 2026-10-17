use vybe_compiler::primitives::{self, polyfills};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/qp_integer_before.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("qp-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = if fast {
        polyfills::build_runtime_helpers_for_exports(
            &mut root,
            &["__stdlib_encoding_quoted_printable_decode"],
        )
        .chunks
        .remove(0)
    } else {
        let mut c = Chunk::new("entry");
        c.arity = 1;
        c.local_count = 1;
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        let mut cs = vec![c];
        before::emit_quoted_printable_decode(&mut cs, 0, 1, 0);
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
fn digit(u: u16) -> Option<u16> {
    match u {
        48..=57 => Some(u - 48),
        65..=70 => Some(u - 55),
        97..=102 => Some(u - 87),
        _ => None,
    }
}
fn oracle(s: &str) -> String {
    let u: Vec<_> = s.encode_utf16().collect();
    let mut out = String::new();
    let mut i = 0;
    while i < u.len() {
        if u[i] == 61 {
            let mut look = i + 1;
            while look < u.len() && (u[look] == 32 || u[look] == 9) {
                look += 1;
            }
            if look == u.len() {
                break;
            }
            if u[look] == 13 || u[look] == 10 {
                let first = u[look];
                look += 1;
                if first == 13 && look < u.len() && u[look] == 10 {
                    look += 1;
                }
                i = look;
                continue;
            }
            if look == i + 1 && i + 2 < u.len() {
                if let (Some(a), Some(b)) = (digit(u[i + 1]), digit(u[i + 2])) {
                    out.push(char::from_u32((a * 16 + b) as u32).unwrap());
                    i += 3;
                    continue;
                }
            }
        }
        out.push_str(&String::from_utf16_lossy(&[u[i]]));
        i += 1;
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
fn exhaustive_digits_and_sizes() {
    let mut ws = [worker(false), worker(true)];
    let mut checks = 0;
    let mut domain = 0;
    for base in (0..65536u32).step_by(256) {
        let mut s = String::new();
        for u in base..base + 256 {
            if let Some(ch) = char::from_u32(u) {
                s.push('=');
                s.push(ch);
                s.push_str("0|=0");
                s.push(ch);
                s.push('|');
                domain += 1;
            }
        }
        checks += check(&mut ws, &s);
    }
    let alphabet = ['=', '0', 'A', 'f', 'G', '\r', '\n', ' ', '\t', 'é'];
    for len in 0..5 {
        for mut bits in 0..10usize.pow(len) {
            let s: String = (0..len)
                .map(|_| {
                    let ch = alphabet[bits % 10];
                    bits /= 10;
                    ch
                })
                .collect();
            checks += check(&mut ws, &s);
        }
    }
    let hex = "0123456789ABCDEFabcdef";
    for a in hex.chars() {
        for b in hex.chars() {
            checks += check(&mut ws, &format!("={a}{b}"));
        }
    }
    for s in [
        "",
        "=",
        "= ",
        "=\t\r\nx",
        "abc=\rnext",
        "==41",
        "=00=FF=7f",
        "é😀=41😀é",
        "x\0=00\r\n",
        "=0",
        "=GG=G0=0G",
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
        native_export::export_chunks(&cs, &format!("qp-integer-{fast}"));
        sizes.push(serde_json::json!({"body":cs[1].code.len(),"locals":cs[1].local_count,"native":bytes.len()}));
    }
    assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
    assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
    assert_eq!(sizes[1]["locals"], sizes[0]["locals"]);
    let result =
        serde_json::json!({"checks":checks,"digit_position_scalar_domain":domain,"sizes":sizes});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../QP_INTEGER_GATES.json"),
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
        ("control", "=aF".repeat(256), true),
        ("empty", String::new(), false),
        ("short", "=aF".into(), false),
        ("decimal256", "=09".repeat(256), false),
        ("alpha256", "=aF".repeat(256), false),
        ("plain256", "abcd".repeat(64), false),
        ("malformed256", "=Gg".repeat(256), false),
        ("mixed", "Abc=00=7f=Fa é😀= \t\r\n!".repeat(64), false),
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
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../QP_INTEGER_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
