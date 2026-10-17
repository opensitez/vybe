//! Bulk-memory correctness and paired kernel measurements.
use super::*;

fn memory_bytes(vm: &VM, index: usize) -> Vec<u8> {
    if index == 0 { vm.memory.with_buffer(|bytes| bytes.to_vec()) }
    else { vm.extra_memories[index - 1].with_buffer(|bytes| bytes.to_vec()) }
}

#[test]
fn extra_memory_handle_observes_core_writes_and_snapshot_restore() {
    let mut vm = VM::new();
    vm.extra_memories = vec![crate::shared_memory::SharedMemory::from_vec(vec![0; 16])];
    let handle = vm.extra_memories[0].clone();
    let baseline = vm.snapshot();
    vm.write_memory_bytes(1, 4, &[1, 2, 3]).unwrap();
    assert_eq!(handle.with_buffer(|bytes| bytes[4..7].to_vec()), vec![1, 2, 3]);
    assert_eq!(vm.read_memory_array::<3>(1, 4).unwrap(), [1, 2, 3]);
    vm.reset_to(&baseline);
    assert_eq!(handle.with_buffer(|bytes| bytes[4..7].to_vec()), vec![0; 3]);
}

#[test]
fn copy_is_memmove_and_traps_before_mutating() {
    for shared in [false, true] {
        for (dst, src, count) in [(7, 0, 20), (0, 7, 20), (5, 5, 20), (32, 32, 0)] {
            let mut vm = VM::new();
            let seed: Vec<u8> = (0..32).collect();
            vm.memory = crate::shared_memory::SharedMemory::from_vec(seed.clone());
            let peer = shared.then(|| vm.memory.clone());
            let mut expected = seed;
            expected.copy_within(src..src + count, dst);
            vm.copy_memory_bytes(0, dst, 0, src, count).unwrap();
            assert_eq!(memory_bytes(&vm, 0), expected);
            if let Some(peer) = peer { assert_eq!(peer.with_buffer(|b| b.to_vec()), expected); }
        }
        for (dst, src, count) in [(33, 0, 0), (0, 33, 0), (20, 0, 20), (0, 20, 20), (usize::MAX, 0, 2)] {
            let mut vm = VM::new();
            vm.memory = crate::shared_memory::SharedMemory::from_vec(vec![7; 32]);
            let _peer = shared.then(|| vm.memory.clone());
            assert!(vm.copy_memory_bytes(0, dst, 0, src, count).is_err());
            assert_eq!(memory_bytes(&vm, 0), vec![7; 32]);
        }
    }
}

#[test]
fn cross_memory_copy_preserves_sources_and_exact_ranges() {
    for dst in 0..3 {
        for src in 0..3 {
            let mut vm = VM::new();
            vm.memory = crate::shared_memory::SharedMemory::from_vec((0..64).collect());
            vm.extra_memories = vec![
                crate::shared_memory::SharedMemory::from_vec((64..128).collect()),
                crate::shared_memory::SharedMemory::from_vec((128..192).collect()),
            ];
            let before: Vec<_> = (0..3).map(|index| memory_bytes(&vm, index)).collect();
            let mut expected = before.clone();
            expected[dst][3..20].copy_from_slice(&before[src][7..24]);
            vm.copy_memory_bytes(dst, 3, src, 7, 17).unwrap();
            for index in 0..3 { assert_eq!(memory_bytes(&vm, index), expected[index]); }
        }
    }
}

#[test]
fn init_drop_zero_length_and_overflow_preserve_wasm_bounds() {
    for shared in [false, true] {
        for memidx in 0..2 {
            let mut vm = VM::new();
            vm.memory = crate::shared_memory::SharedMemory::from_vec(vec![9; 32]);
            let _peer = shared.then(|| vm.memory.clone());
            vm.extra_memories = vec![crate::shared_memory::SharedMemory::from_vec(vec![9; 32])];
            vm.data_segments = vec![vec![1, 2, 3, 4]];
            vm.init_memory_bytes(0, memidx, 7, 1, 2).unwrap();
            let mut expected = vec![9; 32];
            expected[7..9].copy_from_slice(&[2, 3]);
            assert_eq!(memory_bytes(&vm, memidx), expected);
            for (dst, src, count) in [(31, 0, 2), (0, 3, 2), (usize::MAX, 0, 1)] {
                assert!(vm.init_memory_bytes(0, memidx, dst, src, count).is_err());
                assert_eq!(memory_bytes(&vm, memidx), expected);
            }
            vm.dropped_data.insert(0);
            vm.init_memory_bytes(0, memidx, 32, 0, 0).unwrap();
            for (dst, src, count) in [(33, 0, 0), (0, 1, 0), (0, 0, 1)] {
                assert!(vm.init_memory_bytes(0, memidx, dst, src, count).is_err());
                assert_eq!(memory_bytes(&vm, memidx), expected);
            }
        }
    }
}

#[test]
fn init_all_small_ranges_validate_both_bounds_before_mutation() {
    for memidx in 0..2 {
        for limit in 0usize..=8 {
            for segment_len in 0usize..=8 {
                for dropped in [false, true] {
                    let mut vm = VM::new();
                    vm.data_segments = vec![(0..segment_len as u8).collect()];
                    if dropped { vm.dropped_data.insert(0); }
                    let available = if dropped { 0 } else { segment_len };
                    for dst in 0..=limit + 1 {
                        for src in 0..=segment_len + 1 {
                            for count in 0..=limit.max(segment_len) + 1 {
                                let seed = vec![0x91; limit];
                                vm.memory = crate::shared_memory::SharedMemory::from_vec(seed.clone());
                                vm.extra_memories = vec![crate::shared_memory::SharedMemory::from_vec(seed.clone())];
                                let valid = dst <= limit && src <= available
                                    && count <= limit - dst && count <= available - src;
                                let result = vm.init_memory_bytes(0, memidx, dst, src, count);
                                assert_eq!(result.is_ok(), valid,
                                    "memory={memidx} limit={limit} segment={segment_len} dropped={dropped} dst={dst} src={src} count={count}");
                                let mut expected = seed.clone();
                                if valid && count != 0 {
                                    expected[dst..dst + count]
                                        .copy_from_slice(&vm.data_segments[0][src..src + count]);
                                }
                                assert_eq!(memory_bytes(&vm, memidx), expected);
                                assert_eq!(memory_bytes(&vm, 1 - memidx), seed);
                            }
                        }
                    }
                }
            }
        }
    }
}

#[test]
fn fill_bounds_and_zero_count_do_not_modify_unrelated_bytes() {
    for shared in [false, true] {
        let mut vm = VM::new();
        vm.memory = crate::shared_memory::SharedMemory::from_vec(vec![9; 32]);
        let _peer = shared.then(|| vm.memory.clone());
        vm.fill_memory_bytes(0, 7, 13, 0xab).unwrap();
        let mut expected = vec![9; 32];
        expected[7..20].fill(0xab);
        vm.fill_memory_bytes(0, 32, 0, 0).unwrap();
        for (dst, count) in [(33, 0), (31, 2), (usize::MAX, 1)] {
            assert!(vm.fill_memory_bytes(0, dst, count, 0).is_err());
            assert_eq!(memory_bytes(&vm, 0), expected);
        }
    }
}

#[test]
fn bulk_operands_keep_independent_unsigned_widths() {
    let mut vm = VM::new();
    vm.stack = vec![Value::I64(7), Value::I32(-1), Value::I32(-1)];
    assert_eq!(vm.pop_bulk_memory_operands(true, false, false).unwrap(),
        (7, u32::MAX as usize, u32::MAX as usize));
    assert!(vm.stack.is_empty());
    vm.stack = vec![Value::I32(-1), Value::I64(9), Value::I32(3)];
    assert_eq!(vm.pop_bulk_memory_operands(false, true, false).unwrap(),
        (u32::MAX as usize, 9, 3));
    vm.stack = vec![Value::I64(7), Value::I64(-1), Value::I64(0)];
    let result = vm.pop_bulk_memory_operands(true, true, true);
    if usize::BITS == 64 { assert_eq!(result.unwrap(), (7, usize::MAX, 0)); }
    else { assert!(result.is_err()); }
}

#[test]
fn combined_copy_guard_matches_all_small_ranges_and_overflow() {
    for limit in 0usize..=24 {
        for src in 0..=limit + 1 {
            for dst in 0..=limit + 1 {
                for count in 0..=limit + 1 {
                    let mut actual: Vec<u8> = (0..limit as u8).collect();
                    let mut expected = actual.clone();
                    let valid = src <= limit && dst <= limit
                        && count <= limit - src && count <= limit - dst;
                    let result = VM::copy_memory_range(&mut actual, dst, src, count);
                    assert_eq!(result.is_ok(), valid, "limit={limit} src={src} dst={dst} count={count}");
                    if valid { expected.copy_within(src..src + count, dst); }
                    assert_eq!(actual, expected);
                }
            }
        }
    }
    for (dst, src, count) in [
        (0, usize::MAX, 0), (usize::MAX, 0, 0),
        (0, usize::MAX, 2), (usize::MAX, 0, 2),
        (1, 1, usize::MAX), (0, 0, usize::MAX),
    ] {
        let mut actual = [7; 32];
        assert!(VM::copy_memory_range(&mut actual, dst, src, count).is_err());
        assert_eq!(actual, [7; 32]);
    }
}

fn previous_copy(vm: &mut VM, dst: usize, src: usize, count: usize) -> Result<(), crate::VMError> {
    if let Some(result) = vm.memory.with_buffer_mut_exclusive(|memory| {
        let src_end = VM::checked_memory_end(src, count, memory.len())?;
        VM::checked_memory_end(dst, count, memory.len())?;
        memory.copy_within(src..src_end, dst);
        Ok(())
    }) { return result; }
    vm.memory.with_buffer_mut(|memory| {
        let src_end = VM::checked_memory_end(src, count, memory.len())?;
        VM::checked_memory_end(dst, count, memory.len())?;
        memory.copy_within(src..src_end, dst);
        Ok(())
    })
}

#[test]
#[ignore = "paired bulk copy helper benchmark, not complete VM or Doom timing"]
fn copy_kernel_benchmark() {
    use std::hint::black_box;
    use std::time::Instant;
    for shared in [false, true] {
        for size in [0usize, 8, 64, 4096, 1 << 20, 16 << 20] {
            for (kind, dst, src) in [("disjoint", size + 32, 3), ("overlap", 7, 3), ("identity", 3, 3)] {
                let seed: Vec<u8> = (0..size * 2 + 64).map(|i| i.wrapping_mul(31) as u8).collect();
                let mut baseline = VM::new();
                let mut current = VM::new();
                baseline.memory = crate::shared_memory::SharedMemory::from_vec(seed.clone());
                current.memory = crate::shared_memory::SharedMemory::from_vec(seed);
                let _baseline_peer = shared.then(|| baseline.memory.clone());
                let _current_peer = shared.then(|| current.memory.clone());
                let iterations = if kind == "identity" || size == 0 { 131_072 }
                    else if size >= 1 << 20 { 16 }
                    else if size >= 4096 { 1024 }
                    else { 131_072 };
                for sample in 0..23 {
                    let mut times = [std::time::Duration::ZERO; 2];
                    for candidate in if sample % 2 == 0 { [false, true] } else { [true, false] } {
                        let start = Instant::now();
                        for _ in 0..iterations {
                            let dst = black_box(dst);
                            let src = black_box(src);
                            let count = black_box(size);
                            if candidate { black_box(current.copy_memory_bytes(0, dst, 0, src, count).unwrap()); }
                            else { black_box(previous_copy(&mut baseline, dst, src, count).unwrap()); }
                        }
                        times[candidate as usize] = start.elapsed();
                    }
                    assert_eq!(memory_bytes(&current, 0), memory_bytes(&baseline, 0));
                    if sample < 2 { continue; }
                    println!("copy shared={shared} size={size} kind={kind} sample={} previous={:?} current={:?} ratio={:.3}", sample - 2,
                        times[0], times[1], times[0].as_secs_f64() / times[1].as_secs_f64());
                }
            }
        }
    }
}

fn previous_init(vm: &mut VM, src: usize, dst: usize, count: usize) {
    let seg_len = if vm.dropped_data.contains(&0) {
        0
    } else {
        vm.data_segments.first().map(|segment| segment.len()).unwrap_or(0)
    };
    assert!(src.saturating_add(count) <= seg_len);
    assert!(dst.saturating_add(count) <= vm.mem_len(0));
    if count == 0 { return; }
    let bytes = &vm.data_segments[0][src..src + count];
    if vm.memory.with_buffer_mut_exclusive(|memory| {
        memory[dst..dst + count].copy_from_slice(bytes);
    }).is_none() {
        vm.memory.with_buffer_mut(|memory| memory[dst..dst + count].copy_from_slice(bytes));
    }
}

#[test]
#[ignore = "paired bulk initialization kernel benchmark, not VM dispatch timing"]
fn init_kernel_benchmark() {
    use std::hint::black_box;
    use std::time::Instant;
    for shared in [false, true] {
        for size in [0usize, 8, 64, 4096, 1 << 20] {
            let mut vm = VM::new();
            vm.memory = crate::shared_memory::SharedMemory::from_vec(vec![0; size + 32]);
            let _peer = shared.then(|| vm.memory.clone());
            vm.data_segments = vec![vec![0xab; size + 16]];
            let iterations = if size >= 4096 { 256 } else { 32_768 };
            for sample in 0..5 {
                let mut times = [std::time::Duration::ZERO; 2];
                for candidate in if sample % 2 == 0 { [false, true] } else { [true, false] } {
                    let start = Instant::now();
                    for _ in 0..iterations {
                        if candidate { black_box(vm.init_memory_bytes(0, 0, 7, 3, size).unwrap()); }
                        else { previous_init(&mut vm, 3, 7, size); }
                    }
                    times[candidate as usize] = start.elapsed();
                }
                assert!(memory_bytes(&vm, 0)[7..7 + size].iter().all(|&byte| byte == 0xab));
                println!("init shared={shared} size={size} sample={sample} previous={:?} current={:?} ratio={:.3}",
                    times[0], times[1], times[0].as_secs_f64() / times[1].as_secs_f64());
            }
        }
    }
}

mod array_init_execution {
    use super::*;
    use crate::{chunk::Chunk, opcode::Op};

    fn op(chunk: &mut Chunk, operation: Op) {
        chunk.code.extend_from_slice(&operation.encode());
    }

    fn integer(chunk: &mut Chunk, mut value: i32) {
        op(chunk, Op::I32_CONST);
        loop {
            let byte = (value & 127) as u8;
            value >>= 7;
            let done = (value == 0 && byte & 64 == 0)
                || (value == -1 && byte & 64 != 0);
            chunk.code.push(if done { byte } else { byte | 128 });
            if done { break; }
        }
    }

    fn local(chunk: &mut Chunk, operation: Op) {
        op(chunk, operation);
        chunk.code.extend_from_slice(&0u16.to_be_bytes());
    }

    fn program(dst: i32, src: i32, count: i32, verify: bool) -> Chunk {
        let mut chunk = Chunk::new("script");
        chunk.local_count = 1;
        let mut data = vec![0xa5, 0x5a];
        for value in [7i32, -2, 42] {
            data.extend_from_slice(&value.to_le_bytes());
        }
        chunk.data_segments.push(data);
        for _ in 0..5 { integer(&mut chunk, 99); }
        op(&mut chunk, Op::ARRAY_NEW_FIXED);
        chunk.code.extend_from_slice(&0u16.to_be_bytes());
        chunk.code.extend_from_slice(&5u16.to_be_bytes());
        local(&mut chunk, Op::LOCAL_SET);
        local(&mut chunk, Op::LOCAL_GET);
        integer(&mut chunk, dst);
        integer(&mut chunk, src);
        integer(&mut chunk, count);
        op(&mut chunk, Op::ARRAY_INIT_DATA);
        chunk.code.extend_from_slice(&0u16.to_be_bytes());
        chunk.code.extend_from_slice(&0u16.to_be_bytes());
        if verify {
            let expected = if count == 0 { [99; 5] } else { [99, 7, -2, 42, 99] };
            for (index, expected) in expected.into_iter().enumerate() {
                // Keep the numerator below the array read. Any incorrect
                // value produces a zero denominator and a guest trap.
                integer(&mut chunk, 1);
                local(&mut chunk, Op::LOCAL_GET);
                integer(&mut chunk, index as i32);
                op(&mut chunk, Op::ARRAY_GET);
                integer(&mut chunk, expected);
                op(&mut chunk, Op::I32_EQ);
                op(&mut chunk, Op::I32_DIV_S);
                op(&mut chunk, Op::DROP);
            }
        }
        op(&mut chunk, Op::RETURN);
        chunk
    }

    #[test]
    fn array_init_data_decodes_unaligned_source_and_preserves_sentinels() {
        for history in [false, true] {
            let mut vm = VM::new();
            vm.record_error_context(history);
            vm.run(vec![program(1, 2, 3, true)]).unwrap();
        }
    }

    #[test]
    fn array_init_data_accepts_zero_length_at_both_ends() {
        let mut vm = VM::new();
        vm.run(vec![program(5, 14, 0, true)]).unwrap();
    }

    #[test]
    fn array_init_data_checks_bounds_before_decoding() {
        for (dst, src, count, message) in [
            (4, 2, 2, "out of bounds array access"),
            (0, 3, 3, "out of bounds memory access"),
            (6, 15, 0, "out of bounds array access"),
            (0, 15, 0, "out of bounds memory access"),
            (4, 3, 3, "out of bounds array access"),
        ] {
            let mut vm = VM::new();
            let error = vm.run(vec![program(dst, src, count, false)]).unwrap_err();
            assert!(error.to_string().contains(message), "{error}");
            // A trapped frame retains its local. Verify no prefix was written.
            let array = vm.stack.iter().find_map(|value| {
                if let Value::Object(object) = value {
                    let object = object.lock().unwrap();
                    if let ObjectKind::Array(values) = &object.kind {
                        return Some(values.clone());
                    }
                }
                None
            }).expect("array local retained at trap");
            assert_eq!(array, vec![Value::I32(99); 5]);
        }
    }
}
