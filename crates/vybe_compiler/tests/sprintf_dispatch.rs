#![cfg(feature = "primitive-benchmarks")]
use vybe_compiler::primitives;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/sprintf_dispatch_before.rs"]
mod before;
#[path = "fixtures/sprintf_dispatch_candidate.rs"]
mod candidate;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("sprintf-dispatch-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut imports = Chunk::new("imports");
    let mut c = if fast {
        primitives::sprintf::build_sprintf(&mut imports)
    } else {
        before::build_sprintf(&mut imports)
    };
    c.name = "entry".into();
    c.param_count = 2;
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
fn array(v: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(
        vybe_runtime::value::Object::new_array(v),
    ))
}
fn output(n: &str, v: &serde_json::Value) {
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join(format!("../../SPRINTF_DISPATCH_{n}.json")),
        serde_json::to_vec_pretty(v).unwrap(),
    )
    .unwrap();
}
#[test]
fn production_matches_measured_candidate() {
    let mut imports = Chunk::new("imports");
    let actual = primitives::sprintf::build_sprintf(&mut imports);
    let expected = candidate::build_sprintf(&mut imports);
    assert_eq!(actual.code, expected.code);
    assert_eq!(actual.constants, expected.constants);
    assert_eq!(actual.local_count, expected.local_count);
    assert_eq!(actual.arity, expected.arity);
    assert_eq!(actual.param_count, expected.param_count);
    assert_eq!(format!("{:?}", actual.imports), format!("{:?}", expected.imports));
    assert_eq!(format!("{:?}", actual.global_imports), format!("{:?}", expected.global_imports));
}
#[test]
fn cases_sizes_and_work() {
    let mut v = vm();
    let fs = [
        v.run(program(false)).unwrap(),
        v.run(program(true)).unwrap(),
    ];
    let mut checks = 0;
    for conv in "sdiufFeEgGxXobcBOq".chars() {
        for flags in ["", "#", "+", " ", "0", "-", "#0", "+ ", " +"] {
            for width in ["", "4", "12"] {
                for precision in ["", ".0", ".3"] {
                    for num in [0.0, -0.0, 0.5, -1.5, 15.75, 255.0, -255.0, 4294967295.0] {
                        let fmt = format!("a%{flags}{width}{precision}{conv}z");
                        let args = [
                            Value::String(fmt.clone().into()),
                            array(vec![Value::F64(num)]),
                        ];
                        let a = v.invoke(&fs[0], &args).unwrap();
                        let b = v.invoke(&fs[1], &args).unwrap();
                        assert_eq!(a, b, "fmt={fmt} num={num}");
                        checks += 1;
                    }
                }
            }
        }
    }
    for fmt in [
        "",
        "plain😀text",
        "%%",
        "%",
        "%#",
        "%s|%d|%X|%o|%b|%g",
        "%2$x/%1$o/%2$X",
        "%08.3X/%#b/%#o",
    ] {
        let args = [
            Value::String(fmt.into()),
            array(vec![
                Value::I32(15),
                Value::I32(255),
                Value::I32(-1),
                Value::I32(0),
                Value::F64(1.25),
                Value::I32(3),
            ]),
        ];
        assert_eq!(
            v.invoke(&fs[0], &args).unwrap(),
            v.invoke(&fs[1], &args).unwrap(),
            "fmt={fmt}"
        );
        checks += 1;
    }
    let mut sizes = vec![];
    for fast in [false, true] {
        let cs = program(fast);
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
        native_export::export_chunks(&cs, &format!("sprintf-dispatch-{fast}"));
        sizes.push(serde_json::json!({"body":cs[1].code.len(),"locals":cs[1].local_count,"native":b.len()}));
    }
    let mut work = vec![];
    for fmt in [
        "%12d", "%12u", "%12.3e", "%#x", "%#X", "%#o", "%#b", "%12.3s", "%12.3f", "%12.3g", "%12c", "%q",
    ] {
        let mut counts = vec![];
        for f in &fs {
            v.record_runtime_perf(true);
            v.invoke(
                f,
                &[Value::String(fmt.into()), array(vec![Value::I32(255)])],
            )
            .unwrap();
            counts.push(v.take_runtime_perf().unwrap().dispatches);
        }
        work.push(serde_json::json!({"fmt":fmt,"dispatches":counts}));
    }
    let r = serde_json::json!({"checks":checks,"sizes":sizes,"work":work});
    output("GATES", &r);
    eprintln!("{r}");
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for (label, fmt, control) in [
        ("control", "%#x", true),
        ("hex", "%#x", false),
        ("hex-upper", "%#X", false),
        ("octal", "%#o", false),
        ("binary", "%#b", false),
        ("decimal", "%12d", false),
        ("unsigned", "%12u", false),
        ("exponential", "%12.3e", false),
        ("char", "%12c", false),
        ("fixed", "%12.3f", false),
        ("string", "%12.3s", false),
        ("general", "%12.3g", false),
    ] {
        let mut v = vm();
        let a = v.run(program(false)).unwrap();
        let b = if control {
            a.clone()
        } else {
            v.run(program(true)).unwrap()
        };
        let fs = [a, b];
        let args = [
            Value::String(fmt.repeat(8).into()),
            array(vec![Value::I32(255); 8]),
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
                    let cpu = thread_cpu_seconds();
                    let t = Instant::now();
                    std::hint::black_box(v.invoke(&fs[side], &args).unwrap());
                    secs[side] += t.elapsed().as_secs_f64();
                    if let (Some(start), Some(end)) = (cpu, thread_cpu_seconds()) {
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
        let mut cpu_ratios: Vec<_> = cpu_samples[0].iter().zip(&cpu_samples[1])
            .filter(|(_, b)| **b > 0.0).map(|(a,b)| a/b).collect();
        cpu_ratios.sort_by(f64::total_cmp);
        let cpu_median = cpu_ratios.get(cpu_ratios.len()/2).copied();
        rows.push(serde_json::json!({"thread_cpu_seconds_per_call":cpu_samples,"median_thread_cpu_speedup":cpu_median,"label":label,"fmt":fmt,"control":control,"median_speedup":ratios[3],"seconds_per_call":samples}));
    }
    let r = serde_json::json!({"pairs":7,"same_vm":true,"equal_calls_interleaved":true,"minimum_sample_ms":500,"cases":rows});
    output("MEASUREMENTS", &r);
    eprintln!("{r}");
}

#[test]
fn standard_br_table_accepts_i32_producers() {
    for producer in 0..5 {
        let mut root = Chunk::new("br-table-selector-owner");
        primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
        root.emit_op(Op::RETURN, 0);
        let mut c = Chunk::new("entry");
        c.arity = 1;
        c.param_count = 1;
        c.local_count = 1;
        let default = c.emit_block(0);
        let one = c.emit_block(0);
        let zero = c.emit_block(0);
        if producer == 0 {
            c.emit_i32_const(1, 0);
        } else {
            c.emit_op_u16(Op::LOCAL_GET, 0, 0);
            match producer {
                1 => {}
                2 => {
                    c.emit_i32_const(1, 0);
                    c.emit_op(Op::I32_ADD, 0);
                }
                3 => {
                    let import = c.add_import("wasm:js-number", "toI32");
                    c.emit_call(import, 1, 0);
                }
                _ => {
                    c.emit_f64_const(0.0, 0);
                    c.emit_op(Op::F64_LT, 0);
                }
            }
        }
        c.emit_br_table(&[0, 1], 2, 0);
        c.emit_end(0);
        c.patch_block(zero);
        c.emit_i32_const(10, 0);
        c.emit_op(Op::RETURN, 0);
        c.emit_end(0);
        c.patch_block(one);
        c.emit_i32_const(11, 0);
        c.emit_op(Op::RETURN, 0);
        c.emit_end(0);
        c.patch_block(default);
        c.emit_i32_const(99, 0);
        c.emit_op(Op::RETURN, 0);
        let mut cs = vec![root, c];
        primitives::globals::declare_free_globals(&mut cs);
        primitives::globals::normalize_global_table(&mut cs);
        std::fs::write(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(format!("../../BR_TABLE_DIAGNOSTIC_{producer}.wasm")),
            vybe_platform_wasm::write_wasm(&cs),
        ).unwrap();
        native_export::export_chunks(&cs, &format!("br-table-selector-{producer}"));
        let mut v = vm();
        let f = v.run(cs).unwrap();
        for x in [-2i32, -1, 0, 1, 2, i32::MAX] {
            let index = match producer {
                0 => 1,
                2 => x.wrapping_add(1),
                4 => i32::from(x < 0),
                _ => x,
            };
            let expected = match index {
                0 => 10,
                1 => 11,
                _ => 99,
            };
            assert_eq!(
                v.invoke(&f, &[Value::I32(x)]).unwrap(),
                Value::I32(expected)
            );
        }
    }
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
