//! Runtime-local assignment and frame-allocation controls.
//! Timing includes VM::run setup and decoding, not compilation or chunk cloning.
//! These are internal bytecode workloads, not a WASM conformance suite.
use std::hint::black_box;
use std::time::Instant;
use vybe_runtime::{Chunk, Op, Value, VM};

#[derive(Clone, Copy, Debug)]
enum Work {
    FrameOnly,
    Store,
    Copy,
    TeeDrop,
}

fn local(chunk: &mut Chunk, op: Op, slot: u16) {
    chunk.emit_op(op, 1);
    chunk.code.extend_from_slice(&slot.to_be_bytes());
    chunk.lines.resize(chunk.code.len(), 1);
}

fn workload(work: Work, locals: u16, repetitions: usize) -> Chunk {
    assert!(locals >= 2);
    let mut chunk = Chunk::new("assignment-allocation-control");
    chunk.local_count = locals;
    chunk.scratch_high_water = locals;
    chunk.result_arity = 1;
    let target = locals - 1;
    chunk.emit_i32_const(42, 1);
    local(&mut chunk, Op::LOCAL_SET, 0);
    for _ in 0..repetitions {
        match work {
            Work::FrameOnly => {}
            Work::Store => {
                chunk.emit_i32_const(42, 1);
                local(&mut chunk, Op::LOCAL_SET, target);
            }
            Work::Copy => {
                local(&mut chunk, Op::LOCAL_GET, 0);
                local(&mut chunk, Op::LOCAL_SET, target);
            }
            Work::TeeDrop => {
                local(&mut chunk, Op::LOCAL_GET, 0);
                local(&mut chunk, Op::LOCAL_TEE, target);
                chunk.emit_op(Op::DROP, 1);
            }
        }
    }
    let result_slot = if matches!(work, Work::FrameOnly) || repetitions == 0 {
        0
    } else {
        target
    };
    local(&mut chunk, Op::LOCAL_GET, result_slot);
    chunk.emit_op(Op::RETURN, 1);
    chunk
}

fn execute(chunk: Chunk) -> f64 {
    let mut vm = VM::new();
    let chunks = vec![chunk];
    let start = Instant::now();
    let result = vm.run(black_box(chunks)).expect("assignment workload must run");
    let elapsed = start.elapsed().as_secs_f64();
    assert!(matches!(result, Value::I32(42)));
    elapsed
}

#[test]
fn assignment_controls_preserve_results_across_frame_sizes() {
    for locals in [2, 8, 256, 4096] {
        for work in [Work::FrameOnly, Work::Store, Work::Copy, Work::TeeDrop] {
            execute(workload(work, locals, 32));
        }
    }
}

#[test]
#[ignore = "runtime allocation/assignment baseline; not a before/after speedup claim"]
fn assignment_and_frame_allocation_loaded_benchmark() {
    let cases = [
        (Work::FrameOnly, 8, 0),
        (Work::FrameOnly, 256, 0),
        (Work::FrameOnly, 4096, 0),
        (Work::Store, 8, 8192),
        (Work::Copy, 8, 8192),
        (Work::TeeDrop, 8, 8192),
    ];
    let chunks: Vec<_> = cases.iter().map(|&(work, locals, count)| {
        workload(work, locals, count)
    }).collect();
    let mut samples: Vec<Vec<f64>> = (0..cases.len()).map(|_| Vec::new()).collect();
    // Rotate order to avoid always exposing one case to the same load phase.
    for round in 0..23 {
        for step in 0..cases.len() {
            let index = (round + step) % cases.len();
            let elapsed = execute(chunks[index].clone());
            if round >= 2 { samples[index].push(elapsed); }
        }
    }
    for (index, &(work, locals, count)) in cases.iter().enumerate() {
        samples[index].sort_by(f64::total_cmp);
        eprintln!(
            "assignment-allocation work={work:?} locals={locals} repetitions={count} bytes={} p50_ms={:.6} p95_ms={:.6}; includes VM::run setup/decode",
            chunks[index].code.len(), samples[index][10] * 1000.0,
            samples[index][19] * 1000.0,
        );
    }
}
