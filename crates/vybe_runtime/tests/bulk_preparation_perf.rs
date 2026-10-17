//! Separate context-free preparation from linked bulk-operation execution.
//! These timings still include VM linking, validation, and execution setup.
use std::time::{Duration, Instant};
use vybe_runtime::{Chunk, Op, PreparedLinkedCode, Value, VM};

fn constant(chunk: &mut Chunk, mut value: i32) {
    chunk.code.extend_from_slice(&Op::I32_CONST.encode());
    loop {
        let byte = (value & 0x7f) as u8;
        value >>= 7;
        let done = (value == 0 && byte & 0x40 == 0)
            || (value == -1 && byte & 0x40 != 0);
        chunk.code.push(byte | if done { 0 } else { 0x80 });
        if done {
            break;
        }
    }
}

fn batch(size: i32, operations: usize) -> Vec<Chunk> {
    let mut chunk = Chunk::new("bulk preparation benchmark");
    chunk.memory_min_pages = vec![1];
    for _ in 0..operations {
        constant(&mut chunk, 16384);
        constant(&mut chunk, 0x91);
        constant(&mut chunk, size);
        chunk.code.extend_from_slice(&Op::MEMORY_FILL.encode());
        chunk.code.push(0); // memory index
    }
    constant(&mut chunk, 42);
    chunk.code.extend_from_slice(&Op::RETURN.encode());
    vec![chunk]
}

fn prepare(chunks: Vec<Chunk>) -> PreparedLinkedCode {
    match PreparedLinkedCode::prepare(chunks, 32 * 1024 * 1024) {
        Ok(code) => code,
        Err(_) => panic!("small benchmark unexpectedly exceeded preparation budget"),
    }
}

fn check(result: Value) {
    assert!(matches!(result, Value::I32(42)), "unexpected result: {result:?}");
}

fn initialized_vm() -> VM {
    let mut vm = VM::new();
    check(vm.run(batch(0, 0)).unwrap());
    vm
}

fn loaded_batch(size: i32, operations: usize) -> (VM, Value) {
    let mut main = Chunk::new("loaded bulk benchmark");
    main.memory_min_pages = vec![1];
    main.emit_op_u16(Op::REF_FUNC, 1, 0);
    main.emit(0, 0); // no captured upvalues
    main.emit_op(Op::RETURN, 0);
    let function = batch(size, operations).remove(0);
    let mut vm = VM::new();
    let callable = vm.run(vec![main, function]).unwrap();
    (vm, callable)
}

#[test]
fn loaded_bulk_batch_can_be_repeated_without_relinking() {
    for size in [0, 8, 64, 4096] {
        let (mut vm, callable) = loaded_batch(size, 16);
        for _ in 0..8 {
            check(vm.invoke(&callable, &[]).unwrap());
        }
    }
}

#[test]
fn prepared_and_raw_bulk_batches_agree() {
    for size in [0, 8, 64, 4096] {
        let chunks = batch(size, 16);
        let mut raw_vm = initialized_vm();
        check(raw_vm.run_linked_nested(chunks.clone(), vec![]).unwrap());
        let mut prepared_vm = initialized_vm();
        check(prepared_vm.run_linked_nested_prepared(prepare(chunks), vec![]).unwrap());
    }
}

fn median(samples: &mut [Duration]) -> f64 {
    samples.sort_unstable();
    samples[samples.len() / 2].as_secs_f64() * 1e6
}

#[test]
#[ignore = "timing experiment; run explicitly on the target machine"]
fn bulk_preparation_cost_benchmark() {
    for size in [0, 8, 64, 4096] {
        let chunks = batch(size, 4096);
        let mut raw = Vec::with_capacity(15);
        let mut prepared = Vec::with_capacity(15);
        let mut preparation = Vec::with_capacity(15);
        for sample in 0..15 {
            // Alternate ordering to reduce warmup and machine-load bias.
            for prepared_first in [sample % 2 == 0, sample % 2 != 0] {
                let input = chunks.clone();
                let mut vm = initialized_vm();
                vm.record_error_context(false);
                if prepared_first {
                    let start = Instant::now();
                    let code = prepare(input);
                    preparation.push(start.elapsed());
                    let start = Instant::now();
                    let result = vm.run_linked_nested_prepared(code, vec![]).unwrap();
                    prepared.push(start.elapsed());
                    check(result);
                } else {
                    let start = Instant::now();
                    let result = vm.run_linked_nested(input, vec![]).unwrap();
                    raw.push(start.elapsed());
                    check(result);
                }
            }
        }
        println!(
            "fill size={size:4} ops=4096 raw_link_execute_us={:.1} prepared_link_execute_us={:.1} preparation_us={:.1}",
            median(&mut raw), median(&mut prepared), median(&mut preparation),
        );
    }
}

#[test]
#[ignore = "timing experiment; run explicitly on the target machine"]
fn loaded_bulk_execution_cost_benchmark() {
    for history in [false, true] {
        for operations in [1, 16, 4096] {
            for size in [0, 8, 64, 4096] {
                let (mut vm, callable) = loaded_batch(size, operations);
                vm.record_error_context(history);
                check(vm.invoke(&callable, &[]).unwrap());
                let mut samples = Vec::with_capacity(15);
                for _ in 0..15 {
                    let start = Instant::now();
                    for _ in 0..8 {
                        check(vm.invoke(&callable, &[]).unwrap());
                    }
                    samples.push(start.elapsed());
                }
                let invocation_us = median(&mut samples) / 8.0;
                println!(
                    "loaded fill history={history} size={size:4} ops={operations:4} invocation_us={invocation_us:.3} amortized_ns_per_bulk_op={:.1}",
                    invocation_us * 1000.0 / operations as f64,
                );
            }
        }
    }
}
