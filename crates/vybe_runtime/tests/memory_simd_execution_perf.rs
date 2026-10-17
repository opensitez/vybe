//! Prepared bytecode execution gates. Chunk construction and VM creation are
//! outside timing; VM::run may still include module setup and decoding.
//! These measurements are not before/after or whole-application speedups.

use vybe_runtime::{Chunk, Op, VM};
use vybe_runtime::value::Value;

const SEED: [u8; 16] = [0, 1, 2, 127, 128, 254, 255, 64, 3, 9, 17, 31, 63, 129, 191, 253];
const RHS: [u8; 16] = [1, 255, 127, 128, 129, 2, 3, 64, 17, 31, 63, 127, 191, 255, 7, 9];

fn vector(chunk: &mut Chunk, value: [u8; 16]) {
    chunk.emit_op(Op::V128_CONST, 0);
    for byte in value {
        chunk.emit(byte, 0);
    }
}

fn grow_one_page(chunk: &mut Chunk) {
    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_leb_u32(0, 0);
    chunk.emit_op(Op::DROP, 0);
}

fn saturation_program(iterations: usize, memory: bool) -> Chunk {
    let mut chunk = Chunk::new("simd-saturation-execution");
    if memory {
        grow_one_page(&mut chunk);
        chunk.emit_i32_const(0, 0);
        vector(&mut chunk, SEED);
        chunk.emit_op(Op::V128_STORE, 0);
    } else {
        vector(&mut chunk, SEED);
    }
    for _ in 0..iterations {
        if memory {
            // Store address precedes the loaded vector on the operand stack.
            chunk.emit_i32_const(0, 0);
            chunk.emit_i32_const(0, 0);
            chunk.emit_op(Op::V128_LOAD, 0);
        }
        vector(&mut chunk, RHS);
        chunk.emit_op(Op::I8X16_ADD_SAT_U, 0);
        if memory {
            chunk.emit_op(Op::V128_STORE, 0);
        }
    }
    if memory {
        chunk.emit_i32_const(0, 0);
        chunk.emit_op(Op::V128_LOAD, 0);
    }
    chunk.emit_op(Op::RETURN, 0);
    chunk
}

fn saturation_expected(iterations: usize) -> [u8; 16] {
    let mut out = SEED;
    for _ in 0..iterations {
        for (value, rhs) in out.iter_mut().zip(RHS) {
            *value = value.saturating_add(rhs);
        }
    }
    out
}

fn copy_program(iterations: usize) -> Chunk {
    let mut chunk = Chunk::new("overlapping-memory-copy-execution");
    grow_one_page(&mut chunk);
    chunk.emit_i32_const(0, 0);
    vector(&mut chunk, SEED);
    chunk.emit_op(Op::V128_STORE, 0);
    for _ in 0..iterations {
        chunk.emit_i32_const(1, 0);
        chunk.emit_i32_const(0, 0);
        chunk.emit_i32_const(64, 0);
        chunk.emit_op(Op::MEMORY_COPY, 0);
        chunk.emit_leb_u32(0, 0);
        chunk.emit_leb_u32(0, 0);
    }
    chunk.emit_i32_const(49, 0);
    chunk.emit_op(Op::V128_LOAD, 0);
    chunk.emit_op(Op::RETURN, 0);
    chunk
}

fn copy_expected(iterations: usize) -> [u8; 16] {
    let mut memory = [0u8; 65];
    memory[..16].copy_from_slice(&SEED);
    for _ in 0..iterations {
        memory.copy_within(0..64, 1);
    }
    memory[49..65].try_into().unwrap()
}

fn execute(chunk: Chunk, expected: [u8; 16]) {
    let value = VM::new().run(vec![chunk]).expect("bytecode execution failed");
    match value {
        Value::V128(actual) => assert_eq!(actual, expected),
        other => panic!("expected vector, got {other:?}"),
    }
}

#[test]
fn saturation_dispatch_matches_lane_reference() {
    for count in [0, 1, 2, 17, 257] {
        execute(saturation_program(count, false), saturation_expected(count));
    }
}

#[test]
fn vector_memory_roundtrip_matches_lane_reference() {
    for count in [0, 1, 2, 17, 257] {
        execute(saturation_program(count, true), saturation_expected(count));
    }
}

#[test]
fn overlapping_copy_dispatch_matches_memmove_reference() {
    for count in [0, 1, 2, 17, 64, 257] {
        execute(copy_program(count), copy_expected(count));
    }
}

#[test]
#[ignore = "prepared bytecode module-run performance measurement"]
fn prepared_memory_simd_execution_benchmark() {
    use std::time::Instant;
    const ITERATIONS: usize = 20_000;
    fn measure(chunk: Chunk, expected: [u8; 16]) -> f64 {
        let mut vm = VM::new();
        let modules = vec![chunk];
        let start = Instant::now();
        let value = vm.run(modules).expect("benchmark execution failed");
        let elapsed = start.elapsed().as_secs_f64();
        match value {
            Value::V128(actual) => assert_eq!(actual, expected),
            other => panic!("expected vector, got {other:?}"),
        }
        elapsed
    }
    let mut samples = [[0.0f64; 5]; 3];
    let saturation = saturation_expected(ITERATIONS);
    let copied = copy_expected(ITERATIONS);
    // Alternate workload order to reduce systematic first/last bias.
    for pair in 0..5 {
        let order = if pair % 2 == 0 { [0, 1, 2] } else { [2, 1, 0] };
        for case in order {
            samples[case][pair] = match case {
                0 => measure(saturation_program(ITERATIONS, false), saturation),
                1 => measure(saturation_program(ITERATIONS, true), saturation),
                _ => measure(copy_program(ITERATIONS), copied),
            };
        }
    }
    for (name, mut times) in ["register-saturation", "load-saturate-store", "overlap-copy-64"]
        .into_iter().zip(samples)
    {
        times.sort_by(f64::total_cmp);
        eprintln!("{name}: median {:.3} ms, {:.1} ns/iteration, samples {times:?}; includes VM::run setup/decode",
            times[2] * 1000.0, times[2] * 1e9 / ITERATIONS as f64);
    }
}
