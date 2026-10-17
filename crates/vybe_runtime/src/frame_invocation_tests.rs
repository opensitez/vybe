//! Loaded call execution: compilation, decoding, and loading are off timer.
use crate::{Chunk, Op, Value, VM};
use std::{hint::black_box, time::Instant};

fn loaded(arity: u8, locals: u16, history: bool) -> (VM, Value) {
    let mut entry = Chunk::new("frame-benchmark-loader");
    entry.emit_op_u16(Op::REF_FUNC, 1, 1);
    entry.code.push(0); // No captured upvalues; the count is one byte.
    entry.emit_op(Op::RETURN, 1);
    entry.lines.resize(entry.code.len(), 1);
    let mut body = Chunk::new("frame-benchmark-body");
    body.arity = arity;
    body.local_count = locals;
    body.scratch_high_water = locals;
    body.result_arity = 1;
    body.emit_i32_const(42, 1);
    body.emit_op(Op::RETURN, 1);
    let mut vm = VM::new();
    vm.record_error_context(history);
    let callable = vm.run(vec![entry, body]).unwrap();
    (vm, callable)
}

#[test]
fn loaded_frame_calls_preserve_results_across_argument_and_local_sizes() {
    for (arity, argc) in [(0u8, 0usize), (2, 0), (2, 2), (2, 4)] {
        for locals in [2u16, 16, 128, 1024] {
            for history in [false, true] {
                let (mut vm, callable) = loaded(arity, locals, history);
                let args: Vec<Value> = (0..argc).map(|i| Value::I32(i as i32 + 1)).collect();
                for _ in 0..17 {
                    assert_eq!(vm.invoke(&callable, &args).unwrap(), Value::I32(42));
                }
            }
        }
    }
}

#[test]
#[ignore = "loaded invocation/frame-growth timing benchmark"]
fn loaded_frame_invocation_benchmark() {
    for history in [false, true] {
        for (arity, argc) in [(0u8, 0usize), (2, 0), (2, 2), (2, 4)] {
            let mut rows: Vec<_> = [2u16, 16, 128, 1024].into_iter().map(|locals| {
                let (vm, callable) = loaded(arity, locals, history);
                (locals, vm, callable, Vec::<f64>::new())
            }).collect();
            let args: Vec<Value> = (0..argc).map(|i| Value::I32(i as i32 + 1)).collect();
            // Prime frame storage/caches before any measurement.
            for (_, vm, callable, _) in &mut rows {
                for _ in 0..32 { assert_eq!(vm.invoke(callable, &args).unwrap(), Value::I32(42)); }
            }
            for trial in 0..31 {
                for step in 0..rows.len() {
                    let index = (trial + step) % rows.len();
                    let (_, vm, callable, samples) = &mut rows[index];
                    let start = Instant::now();
                    for _ in 0..1024 {
                        assert_eq!(black_box(vm.invoke(black_box(callable), black_box(&args)).unwrap()), Value::I32(42));
                    }
                    samples.push(start.elapsed().as_secs_f64() * 1e9 / 1024.0);
                }
            }
            for (locals, _, _, samples) in &mut rows {
                samples.sort_by(f64::total_cmp);
                println!("loaded-frame history={history} arity={arity} argc={argc} locals={locals} p50_ns={:.2} p95_ns={:.2}",
                    samples[15], samples[29]);
            }
        }
    }
}

#[test]
#[ignore = "bounded large-frame profiler workload; run explicitly"]
fn loaded_large_frame_profile_workload() {
    let (mut vm, callable) = loaded(0, 1024, true);
    for _ in 0..32 { assert_eq!(vm.invoke(&callable, &[]).unwrap(), Value::I32(42)); }
    let start = Instant::now();
    let mut calls = 0usize;
    while start.elapsed() < std::time::Duration::from_secs(10) {
        for _ in 0..4096 {
            assert_eq!(black_box(vm.invoke(black_box(&callable), &[]).unwrap()), Value::I32(42));
        }
        calls += 4096;
    }
    println!("loaded-large-frame-profile calls={calls} elapsed_s={:.3}", start.elapsed().as_secs_f64());
}
