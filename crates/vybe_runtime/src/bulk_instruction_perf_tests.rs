//! Complete bulk-instruction batches; timings include VM setup/predecode.
use super::*;
use std::{hint::black_box, time::Instant};

fn constant(chunk: &mut Chunk, mut value: i32) {
    chunk.code.extend_from_slice(&Op::I32_CONST.encode());
    loop {
        let byte = (value & 127) as u8;
        value >>= 7;
        let done = (value == 0 && byte & 64 == 0) || (value == -1 && byte & 64 != 0);
        chunk.code.push(if done { byte } else { byte | 128 });
        if done { break; }
    }
}

fn program(size: usize, operations: usize, fill: bool) -> Chunk {
    let mut chunk = Chunk::new("script");
    chunk.memory_min_pages = vec![1];
    constant(&mut chunk, 0);
    constant(&mut chunk, 0x91);
    constant(&mut chunk, size as i32);
    chunk.code.extend_from_slice(&Op::MEMORY_FILL.encode());
    chunk.code.push(0);
    for _ in 0..operations {
        constant(&mut chunk, 16384);
        constant(&mut chunk, if fill { 0x91 } else { 0 });
        constant(&mut chunk, size as i32);
        chunk.code.extend_from_slice(&if fill { Op::MEMORY_FILL } else { Op::MEMORY_COPY }.encode());
        chunk.code.push(0);
        if !fill { chunk.code.push(0); }
    }
    constant(&mut chunk, 0);
    chunk.code.extend_from_slice(&Op::RETURN.encode());
    chunk
}

fn validate(vm: &VM, size: usize) {
    vm.memory.with_buffer(|bytes| {
        assert!(bytes[..size].iter().all(|byte| *byte == 0x91));
        assert!(bytes[16384..16384 + size].iter().all(|byte| *byte == 0x91));
        assert_eq!(bytes[16384 + size], 0);
    });
}

fn warmed_program(size: usize, operations: usize, fill: bool, history: bool) -> (VM, Value) {
    let mut body = program(size, operations, fill);
    body.name = "bulk_batch".to_owned();
    body.result_arity = 1;
    body.func_sig = Some((String::new(), "i32".to_owned()));
    let mut script = Chunk::new("<script>");
    script.memory_min_pages = std::mem::take(&mut body.memory_min_pages);
    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.code.push(0); // No captured upvalues.
    script.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    vm.record_error_context(history);
    let callable = vm.run(vec![script, body]).unwrap();
    assert_eq!(vm.invoke(&callable, &[]).unwrap(), Value::I32(0));
    validate(&vm, size);
    (vm, callable)
}

#[test]
fn warmed_bulk_batches_match_cold_payloads_and_repeated_invocations() {
    for history in [false, true] {
        for fill in [false, true] {
            for size in [0, 8, 64, 4096] {
                let mut cold = VM::new();
                cold.record_error_context(history);
                assert_eq!(cold.run(vec![program(size, 16, fill)]).unwrap(), Value::I32(0));
                let expected = cold.memory.with_buffer(|bytes| bytes.to_vec());
                let (mut warm, callable) = warmed_program(size, 16, fill, history);
                for _ in 0..3 {
                    assert_eq!(warm.invoke(&callable, &[]).unwrap(), Value::I32(0));
                    warm.memory.with_buffer(|bytes| assert_eq!(bytes, expected.as_slice()));
                }
            }
        }
    }
}

#[test]
#[ignore = "warmed VM bulk-operation control; excludes module preparation, includes invoke/constants"]
fn warmed_bulk_instruction_fixed_cost_benchmark() {
    for history in [false, true] {
        for fill in [false, true] {
            for size in [0usize, 8, 64, 4096] {
                let operations = 4096;
                let (mut vm, callable) = warmed_program(size, operations, fill, history);
                let mut bytes = vec![0u8; 65536];
                bytes[..size].fill(0x91);
                let mut vm_times = Vec::new();
                let mut native_times = Vec::new();
                for sample in 0..11 {
                    for run_vm in if sample % 2 == 0 { [true, false] } else { [false, true] } {
                        let start = Instant::now();
                        if run_vm {
                            assert_eq!(black_box(vm.invoke(&callable, &[]).unwrap()), Value::I32(0));
                            vm_times.push(start.elapsed().as_nanos());
                        } else {
                            for _ in 0..operations {
                                let bytes = black_box(&mut bytes);
                                if fill {
                                    bytes[16384..16384 + size].fill(black_box(0x91));
                                } else {
                                    bytes.copy_within(black_box(0)..size, 16384);
                                }
                            }
                            native_times.push(start.elapsed().as_nanos());
                        }
                    }
                    validate(&vm, size);
                    assert!(bytes[16384..16384 + size].iter().all(|byte| *byte == 0x91));
                    assert_eq!(bytes[16384 + size], 0);
                }
                vm_times.sort_unstable();
                native_times.sort_unstable();
                eprintln!("warmed-bulk history={history} fill={fill} bytes={size} operations={operations} invoke_p50_us={:.1} native_p50_us={:.1}; module preparation excluded, invoke and three constants per operation included",
                    vm_times[5] as f64 / 1000.0, native_times[5] as f64 / 1000.0);
            }
        }
    }
}

#[test]
fn bulk_batches_preserve_exact_payloads_with_history_enabled_or_disabled() {
    for history in [false, true] {
        for fill in [false, true] {
            for size in [0, 8, 64, 4096] {
                let mut vm = VM::new();
                vm.record_error_context(history);
                assert_eq!(vm.run(vec![program(size, 16, fill)]).unwrap(), Value::I32(0));
                validate(&vm, size);
            }
        }
    }
}

#[test]
#[ignore = "bulk instruction/native kernel control; includes VM setup, not pure dispatch"]
fn bulk_instruction_fixed_cost_benchmark() {
    for history in [false, true] {
        for fill in [false, true] {
            for size in [0usize, 8, 64, 4096] {
                let operations = 4096;
                let chunk = program(size, operations, fill);
                let mut vm_times = Vec::new();
                let mut native_times = Vec::new();
                for _ in 0..11 {
                    let mut vm = VM::new();
                    vm.record_error_context(history);
                    let code = chunk.clone();
                    let start = Instant::now();
                    assert_eq!(vm.run(vec![code]).unwrap(), Value::I32(0));
                    vm_times.push(start.elapsed().as_nanos());
                    validate(&vm, size);
                    let mut bytes = vec![0u8; 65536];
                    bytes[..size].fill(0x91);
                    let start = Instant::now();
                    for _ in 0..operations {
                        let bytes = black_box(&mut bytes);
                        if fill {
                            bytes[16384..16384 + size].fill(black_box(0x91));
                        } else {
                            bytes.copy_within(black_box(0)..size, 16384);
                        }
                    }
                    native_times.push(start.elapsed().as_nanos());
                    assert!(bytes[16384..16384 + size].iter().all(|byte| *byte == 0x91));
                }
                vm_times.sort_unstable();
                native_times.sort_unstable();
                eprintln!("bulk-instruction history={history} fill={fill} bytes={size} operations={operations} vm_run_p50_us={:.1} native_p50_us={:.1}; VM setup/predecode and three constants per operation included",
                    vm_times[5] as f64 / 1000.0, native_times[5] as f64 / 1000.0);
            }
        }
    }
}
