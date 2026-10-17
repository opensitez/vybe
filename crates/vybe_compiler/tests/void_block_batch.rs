#![cfg(feature = "primitive-benchmarks")]
use vybe_compiler::primitives;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(n: usize, typed: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("void-block-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("entry");
    c.arity = 1;
    c.param_count = 1;
    c.local_count = 1;
    let labels: Vec<_> = (0..n).map(|_| c.emit_block(0)).collect();
    if typed {
        c.emit_i32_const(7, 0);
        let b = c.emit_block_params(0, 1, 1);
        c.emit_i32_const(1, 0);
        c.emit_op(Op::I32_ADD, 0);
        c.emit_end(0);
        c.patch_block(b);
        c.emit_op(Op::DROP, 0);
    }
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let depths: Vec<_> = (0..n as u32).collect();
    c.emit_br_table(&depths, n as u32 - 1, 0);
    for (i, b) in labels.into_iter().rev().enumerate() {
        c.emit_end(0);
        c.patch_block(b);
        c.emit_i32_const(i as i32, 0);
        c.emit_op(Op::RETURN, 0);
    }
    let mut cs = vec![root, c];
    primitives::globals::declare_free_globals(&mut cs);
    primitives::globals::normalize_global_table(&mut cs);
    cs
}
fn vm(batch: bool, instrumented: bool) -> VM {
    let mut v = VM::new();
    primitives::platforms::register_platforms_all(&mut v);
    v.benchmark_batch_void_blocks(batch);
    v.record_error_context(instrumented);
    v
}
#[test]
fn labels_depths_boundaries_and_instrumentation() {
    let mut checks = 0;
    let mut rows = vec![];
    for n in [2, 3, 31, 32, 33, 64] {
        for typed in [false, true] {
            let cs = program(n, typed);
            native_export::export_chunks(&cs, &format!("void-block-{n}-{typed}"));
            let mut baseline = vm(false, false);
            let a = baseline.run(cs.clone()).unwrap();
            let mut fast = vm(true, false);
            let b = fast.run(cs.clone()).unwrap();
            let mut instrumented = vm(true, true);
            let c = instrumented.run(cs.clone()).unwrap();
            let mut scalar_instrumented = vm(false, true);
            let d = scalar_instrumented.run(cs).unwrap();
            for x in (-1..n as i32).chain([i32::MAX]) {
                let expected = Value::I32(if x >= 0 && (x as usize) < n {
                    x
                } else {
                    n as i32 - 1
                });
                for (v, f) in [
                    (&mut baseline, &a),
                    (&mut fast, &b),
                    (&mut instrumented, &c),
                    (&mut scalar_instrumented, &d),
                ] {
                    assert_eq!(
                        v.invoke(f, &[Value::I32(x)]).unwrap(),
                        expected,
                        "n={n} typed={typed} x={x}"
                    );
                    checks += 1;
                }
            }
            let mut counts = vec![];
            for (v, f) in [
                (&mut baseline, &a),
                (&mut fast, &b),
                (&mut instrumented, &c),
                (&mut scalar_instrumented, &d),
            ] {
                v.record_runtime_perf(true);
                v.invoke(f, &[Value::I32(0)]).unwrap();
                counts.push(v.take_runtime_perf().unwrap().dispatches);
            }
            assert!(counts[1] < counts[0]);
            assert_eq!(counts[2], counts[3]);
            assert_eq!(counts[2], counts[0]);
            rows.push(serde_json::json!({"blocks":n,"typed_boundary":typed,"dispatches":counts}));
        }
    }
    let r = serde_json::json!({"checks":checks,"rows":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../VOID_BLOCK_GATES.json"),
        serde_json::to_vec_pretty(&r).unwrap(),
    )
    .unwrap();
    eprintln!("{r}");
}
#[test]
fn prepared_reload_and_reset_preserve_block_runs() {
    for prepared in [false, true] {
        let mut v = vm(true, false);
        let snap = v.snapshot();
        for _ in 0..2 {
            let mut retained = Vec::new();
            for n in [33, 2, 64] {
                let chunks = program(n, true);
                let f = if prepared {
                    let seal = vybe_runtime::PreparedLinkedCode::prepare(chunks, usize::MAX)
                        .ok().unwrap();
                    v.run_linked_nested_prepared(seal, vec![]).unwrap()
                } else {
                    v.run_linked_nested(chunks, vec![]).unwrap()
                };
                retained.push((n, f));
            }
            for (n, f) in retained {
                for x in [-1, 0, 1, n as i32 - 1, i32::MAX] {
                    let expected = if x >= 0 && x < n as i32 { x } else { n as i32 - 1 };
                    assert_eq!(v.invoke(&f, &[Value::I32(x)]).unwrap(), Value::I32(expected));
                }
                v.record_runtime_perf(true);
                v.invoke(&f, &[Value::I32(0)]).unwrap();
                assert!(v.take_runtime_perf().unwrap().dispatches < n as u64 + 10);
            }
            v.reset_to(&snap);
        }
    }
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for (n, control) in [
        (64, true),
        (2, false),
        (12, false),
        (32, false),
        (64, false),
    ] {
        let cs = program(n, false);
        let mut v = vm(false, false);
        let f = v.run(cs).unwrap();
        for enabled in [false, true] {
            v.benchmark_batch_void_blocks(enabled && !control);
            for _ in 0..16 {
                v.invoke(&f, &[Value::I32(0)]).unwrap();
            }
        }
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            let mut secs = [0.0, 0.0];
            let mut calls = 0;
            let mut turn = pair;
            while secs[0] < 0.5 || secs[1] < 0.5 || calls < 8 {
                for side in if turn % 2 == 0 { [0, 1] } else { [1, 0] } {
                    v.benchmark_batch_void_blocks(side == 1 && !control);
                    let t = Instant::now();
                    std::hint::black_box(v.invoke(&f, &[Value::I32(0)]).unwrap());
                    secs[side] += t.elapsed().as_secs_f64();
                }
                calls += 1;
                turn += 1;
            }
            for side in 0..2 {
                samples[side].push(secs[side] / calls as f64);
            }
        }
        let mut ratios: Vec<_> = samples[0]
            .iter()
            .zip(&samples[1])
            .map(|(a, b)| a / b)
            .collect();
        ratios.sort_by(f64::total_cmp);
        rows.push(serde_json::json!({"blocks":n,"control":control,"median_speedup":ratios[3],"seconds_per_call":samples}));
    }
    let r = serde_json::json!({"same_vm_same_function":true,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../VOID_BLOCK_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&r).unwrap(),
    )
    .unwrap();
    eprintln!("{r}");
}
