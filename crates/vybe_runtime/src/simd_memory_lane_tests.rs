//! Independent byte-level oracles for Wasm SIMD memory instructions.
use crate::{Chunk, Value, VM};
use crate::opcode::Op;
use crate::shared_memory::SharedMemory;

fn leb(chunk: &mut Chunk, mut value: u32) {
    loop {
        let byte = (value & 127) as u8;
        value >>= 7;
        chunk.code.push(byte | if value == 0 { 0 } else { 128 });
        if value == 0 { break; }
    }
}

fn program(op: Op, width: usize, address: i32, offset: u32, memory: usize,
    vector: Option<[u8; 16]>, lane: Option<u8>, store: bool) -> Chunk
{
    let mut chunk = Chunk::new("simd-memory-byte-oracle");
    chunk.emit_i32_const(address, 1);
    if let Some(vector) = vector {
        chunk.emit_op(Op::V128_CONST, 1);
        chunk.code.extend_from_slice(&vector);
    }
    chunk.emit_op(op, 1);
    let align = width.trailing_zeros();
    leb(&mut chunk, align | if memory == 0 { 0x80 } else { 0xc0 });
    leb(&mut chunk, offset);
    if memory != 0 { leb(&mut chunk, memory as u32); }
    if let Some(lane) = lane { chunk.code.push(lane); }
    if store { chunk.emit_i32_const(42, 1); }
    chunk.emit_op(Op::RETURN, 1);
    chunk.lines.resize(chunk.code.len(), 1);
    chunk
}

fn run(chunk: Chunk, memory: usize, aliased: bool) -> (Result<Value, String>, Vec<u8>, Vec<u8>) {
    let mut vm = VM::new();
    let zero: Vec<u8> = (0..256).map(|i| (i * 37 + 11) as u8).collect();
    let one: Vec<u8> = (0..256).map(|i| (i * 19 + 131) as u8).collect();
    vm.memory = SharedMemory::from_vec(zero);
    vm.extra_memories = vec![SharedMemory::from_vec(one)];
    let alias = aliased.then(|| if memory == 0 { vm.memory.clone() } else { vm.extra_memories[0].clone() });
    let result = vm.run(vec![chunk]).map_err(|e| e.to_string());
    let zero = vm.memory.with_buffer(|b| b.to_vec());
    let one = vm.extra_memories[0].with_buffer(|b| b.to_vec());
    if let Some(alias) = alias {
        assert_eq!(alias.with_buffer(|b| b.to_vec()), if memory == 0 { zero.clone() } else { one.clone() });
    }
    (result, zero, one)
}

#[test]
fn lane_memory_operations_preserve_exact_bytes_and_trap_without_partial_store() {
    let vector = std::array::from_fn(|i| (i * 17 + 3) as u8);
    for (load, store, width) in [
        (Op::V128_LOAD8_LANE, Op::V128_STORE8_LANE, 1usize),
        (Op::V128_LOAD16_LANE, Op::V128_STORE16_LANE, 2),
        (Op::V128_LOAD32_LANE, Op::V128_STORE32_LANE, 4),
        (Op::V128_LOAD64_LANE, Op::V128_STORE64_LANE, 8),
    ] {
        for lane in 0..16 / width {
            for (address, offset) in [(0i32, 0u32), (3, 5), (1, 127), (0, 248), (0, 255), (-1, 1)] {
                for memory in [0usize, 1] { for aliased in [false, true] {
                    for is_store in [false, true] {
                        let mut expected_zero: Vec<u8> = (0..256).map(|i| (i * 37 + 11) as u8).collect();
                        let mut expected_one: Vec<u8> = (0..256).map(|i| (i * 19 + 131) as u8).collect();
                        let selected = if memory == 0 { &mut expected_zero } else { &mut expected_one };
                        let effective = address as u32 as u64 + offset as u64;
                        let trap = effective + width as u64 > selected.len() as u64;
                        let mut expected_vector = vector;
                        if !trap {
                            let at = effective as usize;
                            let lane_at = lane * width;
                            if is_store { selected[at..at + width].copy_from_slice(&vector[lane_at..lane_at + width]); }
                            else { expected_vector[lane_at..lane_at + width].copy_from_slice(&selected[at..at + width]); }
                        }
                        let (result, zero, one) = run(program(if is_store { store } else { load }, width,
                            address, offset, memory, Some(vector), Some(lane as u8), is_store), memory, aliased);
                        assert_eq!(result.is_err(), trap,
                            "width={width} lane={lane} address={address} offset={offset} memory={memory} store={is_store}");
                        if !trap { assert_eq!(result.unwrap(), if is_store { Value::I32(42) } else { Value::V128(expected_vector) }); }
                        assert_eq!(zero, expected_zero);
                        assert_eq!(one, expected_one);
                    }
                } }
            }
        }
    }
}

#[test]
fn splat_zero_and_widening_loads_match_scalar_little_endian_oracles() {
    for (op, width) in [
        (Op::V128_LOAD8_SPLAT, 1usize), (Op::V128_LOAD16_SPLAT, 2),
        (Op::V128_LOAD32_SPLAT, 4), (Op::V128_LOAD64_SPLAT, 8),
        (Op::V128_LOAD32_ZERO, 4), (Op::V128_LOAD64_ZERO, 8),
        (Op::V128_LOAD8X8_S, 8), (Op::V128_LOAD8X8_U, 8),
        (Op::V128_LOAD16X4_S, 8), (Op::V128_LOAD16X4_U, 8),
        (Op::V128_LOAD32X2_S, 8), (Op::V128_LOAD32X2_U, 8),
    ] {
        for (address, offset) in [(0i32, 0u32), (1, 3), (0, 127), (0, 248), (0, 255), (-1, 1)] {
            for memory in [0usize, 1] { for aliased in [false, true] {
                let zero: Vec<u8> = (0..256).map(|i| (i * 37 + 11) as u8).collect();
                let one: Vec<u8> = (0..256).map(|i| (i * 19 + 131) as u8).collect();
                let selected = if memory == 0 { &zero } else { &one };
                let effective = address as u32 as u64 + offset as u64;
                let trap = effective + width as u64 > selected.len() as u64;
                let mut expected = [0u8; 16];
                if !trap {
                    let input = &selected[effective as usize..effective as usize + width];
                    match op {
                        Op::V128_LOAD32_ZERO | Op::V128_LOAD64_ZERO => expected[..width].copy_from_slice(input),
                        Op::V128_LOAD8X8_S | Op::V128_LOAD8X8_U => {
                            for (i, &byte) in input.iter().enumerate() {
                                let value = if op == Op::V128_LOAD8X8_S { byte as i8 as i16 } else { byte as i16 };
                                expected[i * 2..i * 2 + 2].copy_from_slice(&value.to_le_bytes());
                            }
                        }
                        Op::V128_LOAD16X4_S | Op::V128_LOAD16X4_U => {
                            for (i, bytes) in input.chunks_exact(2).enumerate() {
                                let raw = u16::from_le_bytes(bytes.try_into().unwrap());
                                let value = if op == Op::V128_LOAD16X4_S { raw as i16 as i32 } else { raw as i32 };
                                expected[i * 4..i * 4 + 4].copy_from_slice(&value.to_le_bytes());
                            }
                        }
                        Op::V128_LOAD32X2_S | Op::V128_LOAD32X2_U => {
                            for (i, bytes) in input.chunks_exact(4).enumerate() {
                                let raw = u32::from_le_bytes(bytes.try_into().unwrap());
                                let value = if op == Op::V128_LOAD32X2_S { raw as i32 as i64 } else { raw as i64 };
                                expected[i * 8..i * 8 + 8].copy_from_slice(&value.to_le_bytes());
                            }
                        }
                        _ => { for lane in expected.chunks_exact_mut(width) { lane.copy_from_slice(input); } }
                    }
                }
                let (result, actual_zero, actual_one) = run(program(op, width, address, offset, memory,
                    None, None, false), memory, aliased);
                assert_eq!(result.is_err(), trap, "op={op:?} address={address} offset={offset} memory={memory}");
                if !trap { assert_eq!(result.unwrap(), Value::V128(expected)); }
                assert_eq!(actual_zero, zero);
                assert_eq!(actual_one, one);
            } }
        }
    }
}
