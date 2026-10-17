#![cfg(feature = "primitive-benchmarks")]
use vybe_compiler::primitives;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/group_cursor_before.rs"]
mod before;
#[path = "fixtures/group_cursor_candidate.rs"]
mod candidate;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool, ascii: bool) -> Vec<Chunk> {
    program_with(
        if fast {
            production
        } else {
            before::emit_group_digits_impl
        },
        ascii,
    )
}
fn production(cs: &mut [Chunk], i: usize, line: u32, ascii: bool) {
    if ascii {
        primitives::strings::emit_group_digits_ascii_numeric(cs, i, line)
    } else {
        primitives::strings::emit_group_digits(cs, i, line)
    }
}
fn program_with(emit: fn(&mut [Chunk], usize, u32, bool), ascii: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("group-cursor-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("entry");
    c.arity = 3;
    c.param_count = 3;
    c.local_count = 3;
    for slot in 0..3 {
        c.emit_op_u16(Op::LOCAL_GET, slot, 0);
    }
    let mut cs = vec![root, c];
    emit(&mut cs, 1, 0, ascii);
    cs[1].emit_op(Op::RETURN, 0);
    primitives::globals::declare_free_globals(&mut cs);
    primitives::globals::normalize_global_table(&mut cs);
    cs
}
fn vm() -> VM {
    let mut v = VM::new();
    primitives::platforms::register_platforms_all(&mut v);
    v
}
fn output(n: &str, v: &serde_json::Value) {
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join(format!("../../GROUP_CURSOR_{n}.json")),
        serde_json::to_vec_pretty(v).unwrap(),
    )
    .unwrap();
}
fn oracle(s: &str, sep: &str, dec: &str, ascii: bool) -> String {
    let u: Vec<u16> = s.encode_utf16().collect();
    let end = u.iter().position(|&x| x == 46).unwrap_or(u.len());
    let start = usize::from(matches!(u.first(), Some(43 | 45)));
    let digits = &u[start.min(end)..end.max(start.min(end))];
    let mut out = String::from_utf16_lossy(&u[..start]);
    for (i, &x) in digits.iter().enumerate() {
        if i > 0 && (digits.len() - i) % 3 == 0 {
            out.push_str(sep);
        }
        if ascii {
            out.push(char::from_u32(x as u32).unwrap());
        } else {
            out.push_str(&String::from_utf16_lossy(&[x]));
        }
    }
    if end < u.len() {
        out.push_str(dec);
        out.push_str(&String::from_utf16_lossy(&u[end + 1..]));
    }
    out
}
#[test]
fn values_sizes_work() {
    let mut checks = 0;
    let mut rows = vec![];
    for ascii in [false, true] {
        let actual = program(true, ascii);
        let measured = program_with(candidate::emit_group_digits_impl, ascii);
        assert_eq!(actual[1].code, measured[1].code);
        assert_eq!(actual[1].imports, measured[1].imports);
        assert_eq!(actual[1].constants, measured[1].constants);
        assert_eq!(actual[1].local_count, measured[1].local_count);
        let mut v = vm();
        let fs = [
            v.run(program(false, ascii)).unwrap(),
            v.run(program(true, ascii)).unwrap(),
        ];
        for n in 0..90 {
            for sign in ["", "+", "-"] {
                for frac in ["", ".", ".75"] {
                    let s = format!("{sign}{}{frac}", "1234567890".repeat(9)[..n].to_string());
                    for (sep, dec) in [("", ":"), (",", "."), ("😀", "界")] {
                        let expected = Value::String(oracle(&s, sep, dec, ascii).into());
                        for f in &fs {
                            assert_eq!(
                                v.invoke(
                                    f,
                                    &[
                                        Value::String(s.clone().into()),
                                        Value::String(sep.into()),
                                        Value::String(dec.into())
                                    ]
                                )
                                .unwrap(),
                                expected,
                                "ascii={ascii} s={s:?}"
                            );
                            checks += 1;
                        }
                    }
                }
            }
        }
        if !ascii {
            for s in ["é界😀123", "😀😀😀", "+é😀123.界😀", "123😀456", "a.b.c"] {
                for f in &fs {
                    assert_eq!(
                        v.invoke(
                            f,
                            &[
                                Value::String(s.into()),
                                Value::String(",".into()),
                                Value::String(".".into())
                            ]
                        )
                        .unwrap(),
                        Value::String(oracle(s, ",", ".", false).into())
                    );
                    checks += 1;
                }
            }
        }
        let mut sizes = vec![];
        for fast in [false, true] {
            let cs = program(fast, ascii);
            let b = vybe_platform_wasm::write_wasm_with_options(
                &cs,
                vybe_platform_wasm::WasmWriteOptions {
                    include_name_section: false,
                    include_vybe_metadata: false,
                },
            );
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&b)
                .unwrap();
            native_export::export_chunks(&cs, &format!("group-cursor-{ascii}-{fast}"));
            sizes.push(serde_json::json!({"body":cs[1].code.len(),"locals":cs[1].local_count,"native":b.len()}));
        }
        assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
        assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
        let mut counts = vec![];
        for f in &fs {
            v.record_runtime_perf(true);
            v.invoke(
                f,
                &[
                    Value::String("1234567890".repeat(10).into()),
                    Value::String(",".into()),
                    Value::String(".".into()),
                ],
            )
            .unwrap();
            counts.push(v.take_runtime_perf().unwrap().dispatches);
        }
        assert!(counts[1] < counts[0]);
        rows.push(serde_json::json!({"ascii":ascii,"sizes":sizes,"dispatches":counts}));
    }
    let r = serde_json::json!({"checks":checks,"rows":rows});
    output("GATES", &r);
    eprintln!("{r}");
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for (ascii, label, s, control) in [
        (false, "control", "1234567890".repeat(10), true),
        (false, "empty", String::new(), false),
        (false, "short", "1234.5".into(), false),
        (false, "long", "1234567890".repeat(10), false),
        (false, "unicode", "é界😀123".repeat(10), false),
        (true, "ascii-short", "1234.5".into(), false),
        (true, "ascii-long", "1234567890".repeat(10), false),
    ] {
        let mut v = vm();
        let a = v.run(program(false, ascii)).unwrap();
        let b = if control {
            a.clone()
        } else {
            v.run(program(true, ascii)).unwrap()
        };
        let fs = [a, b];
        let args = [
            Value::String(s.clone().into()),
            Value::String(",".into()),
            Value::String(".".into()),
        ];
        for f in &fs {
            for _ in 0..8 {
                std::hint::black_box(v.invoke(f, &args).unwrap());
            }
        }
        let mut samples = [vec![], vec![]];
        let mut cpu_samples = [vec![], vec![]];
        for pair in 0..7 {
            let mut secs = [0.0, 0.0];
            let mut cpu_secs = [0.0, 0.0];
            let mut n = 0;
            let mut turn = pair;
            while secs[0] < 0.5 || secs[1] < 0.5 || n < 8 {
                for side in if turn % 2 == 0 { [0, 1] } else { [1, 0] } {
                    let cpu_start = thread_cpu_seconds();
                    let t = Instant::now();
                    std::hint::black_box(v.invoke(&fs[side], &args).unwrap());
                    secs[side] += t.elapsed().as_secs_f64();
                    if let (Some(start), Some(end)) = (cpu_start, thread_cpu_seconds()) {
                        cpu_secs[side] += end - start;
                    }
                }
                n += 1;
                turn += 1;
            }
            for side in 0..2 {
                samples[side].push(secs[side] / n as f64);
                cpu_samples[side].push(cpu_secs[side] / n as f64);
            }
        }
        let mut ratios: Vec<_> = samples[0]
            .iter()
            .zip(&samples[1])
            .map(|(a, b)| a / b)
            .collect();
        ratios.sort_by(f64::total_cmp);
        let mut cpu_ratios: Vec<_> = cpu_samples[0]
            .iter()
            .zip(&cpu_samples[1])
            .filter_map(|(a, b)| (*b > 0.0).then_some(a / b))
            .collect();
        cpu_ratios.sort_by(f64::total_cmp);
        rows.push(serde_json::json!({"ascii":ascii,"label":label,"control":control,"median_speedup":ratios[3],"seconds_per_call":samples,"cpu_seconds_per_call":cpu_samples,"median_cpu_speedup":cpu_ratios.get(cpu_ratios.len()/2)}));
    }
    let r = serde_json::json!({"pairs":7,"same_vm":true,"equal_calls_interleaved":true,"minimum_sample_ms":500,"cases":rows});
    output("MEASUREMENTS", &r);
    eprintln!("{r}");
}

// Diagnostic only: the local macOS SDK defines CLOCK_THREAD_CPUTIME_ID=16.
// Wall time remains recorded; thread CPU time excludes descheduling by other jobs.
#[cfg(target_os = "macos")]
fn thread_cpu_seconds() -> Option<f64> {
    #[repr(C)]
    struct Timespec {
        sec: i64,
        nsec: i64,
    }
    unsafe extern "C" {
        fn clock_gettime(clock: i32, result: *mut Timespec) -> i32;
    }
    let mut result = Timespec { sec: 0, nsec: 0 };
    assert_eq!(unsafe { clock_gettime(16, &mut result) }, 0);
    Some(result.sec as f64 + result.nsec as f64 * 1e-9)
}
#[cfg(not(target_os = "macos"))]
fn thread_cpu_seconds() -> Option<f64> {
    None
}
