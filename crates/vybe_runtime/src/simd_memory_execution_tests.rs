use super::*;
use std::time::Instant;

#[path = "simd_memory_lane_tests.rs"]
mod lane_memory_oracle_tests;

#[derive(Clone, Copy, Debug)]
enum Encoding { Omitted, Marked, Explicit }

fn leb(chunk: &mut Chunk, mut value: u32) {
    loop {
        let mut byte = (value & 127) as u8;
        value >>= 7;
        if value != 0 { byte |= 128; }
        chunk.code.push(byte);
        if value == 0 { break; }
    }
    chunk.lines.resize(chunk.code.len(), 1);
}

fn memarg(chunk: &mut Chunk, encoding: Encoding, offset: u32) {
    if matches!(encoding, Encoding::Omitted) { return; }
    leb(chunk, if matches!(encoding, Encoding::Explicit) { 0xc4 } else { 0x84 });
    leb(chunk, offset);
    if matches!(encoding, Encoding::Explicit) { leb(chunk, 0); }
}

fn local(chunk: &mut Chunk, op: Op) {
    chunk.emit_op(op, 1);
    chunk.code.extend_from_slice(&0u16.to_be_bytes());
    chunk.lines.resize(chunk.code.len(), 1);
}

fn workload(encoding: Encoding, shift: u32, rounds: i32) -> Chunk {
    let mut chunk = Chunk::new("steady-simd-memory");
    chunk.local_count = 8;
    chunk.scratch_high_water = 8;
    chunk.result_arity = 1;
    chunk.emit_i32_const(rounds, 1);
    local(&mut chunk, Op::LOCAL_SET);
    chunk.emit_op_u8_u8(Op::LOOP, 0, 0, 1);
    let address_shift = if matches!(encoding, Encoding::Omitted) { shift as i32 } else { 0 };
    for _ in 0..32 {
        chunk.emit_i32_const(64 + address_shift, 1);
        chunk.emit_i32_const(address_shift, 1);
        chunk.emit_op(Op::V128_LOAD, 1);
        memarg(&mut chunk, encoding, shift);
        chunk.emit_op(Op::V128_STORE, 1);
        memarg(&mut chunk, encoding, shift);
    }
    local(&mut chunk, Op::LOCAL_GET);
    chunk.emit_i32_const(1, 1);
    chunk.emit_op(Op::I32_SUB, 1);
    local(&mut chunk, Op::LOCAL_TEE);
    chunk.emit_op(Op::BR_IF, 1);
    leb(&mut chunk, 0);
    chunk.emit_op(Op::END, 1);
    local(&mut chunk, Op::LOCAL_GET);
    chunk.emit_op_u8_u8(Op::IF, 0, 0, 1);
    chunk.emit_op(Op::UNREACHABLE, 1);
    chunk.emit_op(Op::END, 1);
    chunk.emit_i32_const(42, 1);
    chunk.emit_op(Op::RETURN, 1);
    chunk
}

fn execute(chunk: Chunk, shift: u32, shared: bool) -> f64 {
    let mut vm = VM::new();
    let mut seed = vec![0u8; 65536];
    for (index, byte) in seed[..32].iter_mut().enumerate() { *byte = index as u8 + 1; }
    let start = shift as usize;
    let mut expected = seed.clone();
    expected[64 + start..80 + start].copy_from_slice(&seed[start..16 + start]);
    vm.memory = crate::shared_memory::SharedMemory::from_vec(seed);
    let alias = shared.then(|| vm.memory.clone());
    let chunks = vec![chunk];
    let clock = Instant::now();
    assert_eq!(vm.run(chunks).unwrap(), Value::I32(42));
    let elapsed = clock.elapsed().as_secs_f64();
    assert_eq!(vm.memory.with_buffer(|bytes| bytes.to_vec()), expected);
    if let Some(alias) = alias {
        assert_eq!(alias.with_buffer(|bytes| bytes.to_vec()), expected);
    }
    elapsed
}

#[test]
fn simd_memory_encodings_preserve_unaligned_bytes_and_shared_aliases() {
    for encoding in [Encoding::Omitted, Encoding::Marked, Encoding::Explicit] {
        for shift in [0, 3] {
            for shared in [false, true] { execute(workload(encoding, shift, 17), shift, shared); }
        }
    }
}

#[test]
#[ignore = "steady SIMD memory execution control, not an isolated cache speedup"]
fn loaded_simd_memory_execution_benchmark() {
    let encodings = [Encoding::Omitted, Encoding::Marked, Encoding::Explicit];
    for shared in [false, true] {
        let chunks: Vec<_> = encodings.iter().map(|&encoding| workload(encoding, 3, 8192)).collect();
        let mut samples: Vec<Vec<f64>> = (0..3).map(|_| Vec::new()).collect();
        for round in 0..23 {
            for step in 0..3 {
                let index = (round + step) % 3;
                let elapsed = execute(chunks[index].clone(), 3, shared);
                if round >= 2 { samples[index].push(elapsed); }
            }
        }
        for (index, samples) in samples.iter_mut().enumerate() {
            samples.sort_by(f64::total_cmp);
            eprintln!("simd-memory encoding={:?} shared={shared} pairs=262144 code_bytes={} p50_ms={:.3} p95_ms={:.3}",
                encodings[index], chunks[index].code.len(), samples[10] * 1000.0, samples[19] * 1000.0);
        }
    }
}

fn copy_loop_local(chunk: &mut Chunk, op: Op, slot: u16) {
    chunk.emit_op(op, 1);
    chunk.code.extend_from_slice(&slot.to_be_bytes());
    chunk.lines.resize(chunk.code.len(), 1);
}

fn copy_loop_chunk(src: i32, dst: i32, len: i32, encoding: Encoding) -> Chunk {
    copy_loop_chunk_with_offsets(src, dst, len, encoding, 0, 0)
}

fn copy_loop_chunk_with_offsets(
    src: i32, dst: i32, len: i32, encoding: Encoding,
    load_offset: u32, store_offset: u32,
) -> Chunk {
    let mut chunk = Chunk::new("simd-copy-loop");
    chunk.local_count = 7;
    chunk.scratch_high_water = 7;
    chunk.result_arity = 1;
    for (slot, value) in [(0, src), (1, dst), (2, len), (3, 0)] {
        chunk.emit_i32_const(value, 1);
        copy_loop_local(&mut chunk, Op::LOCAL_SET, slot);
    }
    chunk.emit_op(Op::V128_CONST, 1);
    chunk.code.extend_from_slice(&[0u8; 16]);
    chunk.lines.resize(chunk.code.len(), 1);
    copy_loop_local(&mut chunk, Op::LOCAL_SET, 4);
    chunk.emit_op_u8_u8(Op::BLOCK, 0, 0, 1);
    chunk.emit_op_u8_u8(Op::LOOP, 0, 0, 1);
    copy_loop_local(&mut chunk, Op::LOCAL_GET, 3);
    copy_loop_local(&mut chunk, Op::LOCAL_GET, 2);
    chunk.emit_op(Op::I32_GE_U, 1);
    chunk.emit_op(Op::BR_IF, 1);
    leb(&mut chunk, 1);
    copy_loop_local(&mut chunk, Op::LOCAL_GET, 0);
    copy_loop_local(&mut chunk, Op::LOCAL_GET, 3);
    chunk.emit_op(Op::I32_ADD, 1);
    chunk.emit_op(Op::V128_LOAD, 1);
    memarg(&mut chunk, encoding, load_offset);
    copy_loop_local(&mut chunk, Op::LOCAL_SET, 4);
    copy_loop_local(&mut chunk, Op::LOCAL_GET, 1);
    copy_loop_local(&mut chunk, Op::LOCAL_GET, 3);
    chunk.emit_op(Op::I32_ADD, 1);
    copy_loop_local(&mut chunk, Op::LOCAL_GET, 4);
    chunk.emit_op(Op::V128_STORE, 1);
    memarg(&mut chunk, encoding, store_offset);
    copy_loop_local(&mut chunk, Op::LOCAL_GET, 3);
    chunk.emit_i32_const(16, 1);
    chunk.emit_op(Op::I32_ADD, 1);
    copy_loop_local(&mut chunk, Op::LOCAL_SET, 3);
    chunk.emit_op(Op::BR, 1);
    leb(&mut chunk, 0);
    chunk.emit_op(Op::END, 1);
    chunk.emit_op(Op::END, 1);
    copy_loop_local(&mut chunk, Op::LOCAL_GET, 4);
    chunk.emit_op(Op::RETURN, 1);
    chunk
}

fn run_copy_loop(src: i32, dst: i32, len: i32, encoding: Encoding, shared: bool)
    -> (Result<Value, String>, Vec<u8>)
{
    let mut vm = VM::new();
    vm.memory = crate::shared_memory::SharedMemory::from_vec(
        (0..65536).map(|i| (i % 251) as u8).collect()
    );
    let _alias = shared.then(|| vm.memory.clone());
    let result = vm.run(vec![copy_loop_chunk(src, dst, len, encoding)])
        .map_err(|error| error.to_string());
    let bytes = vm.memory.with_buffer(|bytes| bytes.to_vec());
    (result, bytes)
}

#[test]
fn simd_copy_loops_match_independent_instruction_order_oracle() {
    // Model one load followed by one store per iteration, not memmove.
    // In particular, a later failing store must retain earlier stores.
    for (src, dst) in [
        (0i32, 1i32), (0, 3), (0, 7), (0, 15), (0, 16),
        (0, 17), (17, 0), (3, 1027), (3, 3),
        (65504, 1024), (1024, 65504), (65520, 65528),
        (-16, 1024), (1024, -16),
    ] {
        for len in [0i32, 1, 15, 16, 17, 31, 32, 33, 48, 64, 513] {
            let mut expected: Vec<u8> = (0..65536).map(|i| (i % 251) as u8).collect();
            let mut last = [0u8; 16];
            let mut trap = false;
            let mut counter = 0u32;
            while counter < len as u32 {
                // i32.add wraps, whereas the memory access bounds check does not.
                let load = (src as u32).wrapping_add(counter) as u64;
                let store = (dst as u32).wrapping_add(counter) as u64;
                if load + 16 > expected.len() as u64 { trap = true; break; }
                last.copy_from_slice(&expected[load as usize..load as usize + 16]);
                if store + 16 > expected.len() as u64 { trap = true; break; }
                expected[store as usize..store as usize + 16].copy_from_slice(&last);
                counter = counter.wrapping_add(16);
            }
            for encoding in [Encoding::Omitted, Encoding::Marked, Encoding::Explicit] {
                for shared in [false, true] {
                    let (result, actual) = run_copy_loop(src, dst, len, encoding, shared);
                    assert_eq!(result.is_err(), trap,
                        "src={src} dst={dst} len={len} encoding={encoding:?} shared={shared}");
                    if !trap { assert_eq!(result.unwrap(), Value::V128(last)); }
                    assert_eq!(actual, expected,
                        "src={src} dst={dst} len={len} encoding={encoding:?} shared={shared}");
                }
            }
        }
    }
}

#[test]
fn simd_copy_loop_matches_scalar_bytes_and_last_local() {
    for (src, dst, len) in [
        (3, 1027, 512), (3, 1027, 17), (3, 3, 48),
        (0, 8, 64), (8, 0, 64), (0, 1024, 0),
        (0, 65504, 48), (65504, 0, 48), (-8, 1024, 32),
    ] {
        let reference = run_copy_loop(src, dst, len, Encoding::Marked, true);
        for encoding in [Encoding::Omitted, Encoding::Marked, Encoding::Explicit] {
            for shared in [false, true] {
                let actual = run_copy_loop(src, dst, len, encoding, shared);
                assert_eq!(actual.0.is_err(), reference.0.is_err());
                if reference.0.is_ok() { assert_eq!(actual.0, reference.0); }
                assert_eq!(actual.1, reference.1,
                    "src={src} dst={dst} len={len} encoding={encoding:?} shared={shared}");
            }
        }
    }
}

#[test]
#[ignore = "paired SIMD loop execution comparison; includes VM decode/setup"]
fn simd_copy_loop_execution_benchmark() {
    let mut samples = [Vec::new(), Vec::new()];
    for round in 0..23 {
        for step in 0..2 {
            let index = (round + step) % 2;
            let encoding = Encoding::Marked;
            let chunk = copy_loop_chunk(0, 32768, 32768, encoding);
            let mut vm = VM::new();
            let seed: Vec<u8> = (0..65536).map(|i| (i % 251) as u8).collect();
            let expected = seed[32752..32768].to_vec();
            vm.memory = crate::shared_memory::SharedMemory::from_vec(seed);
            let _alias = (index == 1).then(|| vm.memory.clone());
            let clock = Instant::now();
            let result = vm.run(vec![chunk]).unwrap();
            let elapsed = clock.elapsed().as_secs_f64();
            assert_eq!(result, Value::V128(expected.try_into().unwrap()));
            vm.memory.with_buffer(|bytes| assert_eq!(&bytes[..32768], &bytes[32768..]));
            if round >= 2 { samples[index].push(elapsed); }
        }
    }
    for (index, values) in samples.iter_mut().enumerate() {
        values.sort_by(f64::total_cmp);
        eprintln!("simd-copy-loop path={} vectors=2048 p50_ms={:.3} p95_ms={:.3}",
            if index == 0 { "marked-exclusive-fused" } else { "marked-aliased-scalar" },
            values[10] * 1000.0, values[19] * 1000.0);
    }
}

#[test]
fn late_unaligned_loops_use_compact_independent_cache_buckets() {
    let mut chunk = Chunk::new("late-loop-cache");
    chunk.result_arity = 1;
    for _ in 0..4096 { chunk.emit_op(Op::NOP, 1); }
    // A variable-length immediate makes the subsequent loop starts unaligned.
    chunk.emit_i32_const(128, 1);
    chunk.emit_op(Op::DROP, 1);
    let mut starts = Vec::new();
    for _ in 0..3 {
        chunk.emit_op_u8_u8(Op::BLOCK, 0, 0, 1);
        chunk.emit_op_u8_u8(Op::LOOP, 0, 0, 1);
        starts.push(chunk.code.len());
        chunk.emit_op(Op::NOP, 1);
        chunk.emit_op(Op::END, 1);
        chunk.emit_op(Op::END, 1);
    }
    chunk.emit_i32_const(42, 1);
    chunk.emit_op(Op::RETURN, 1);
    assert!(starts.iter().any(|start| start % 4 != 0));
    let mut vm = VM::new();
    assert_eq!(vm.run(vec![chunk]).unwrap(), Value::I32(42));
    let cache = &vm.loop_superops[0];
    assert_eq!(cache.len(), starts.last().unwrap() / 4 + 1);
    for start in starts { assert_eq!(cache[start / 4], LoopSuperop::None); }
}

#[test]
fn cached_marked_memargs_reuse_the_simd_copy_loop_recognizer() {
    for encoding in [Encoding::Omitted, Encoding::Marked] {
        let mut vm = VM::new();
        vm.memory = crate::shared_memory::SharedMemory::from_vec(vec![0x5a; 65536]);
        assert_eq!(vm.run(vec![copy_loop_chunk(3, 1027, 48, encoding)]).unwrap(),
            Value::V128([0x5a; 16]));
        assert!(vm.loop_superops[0].iter().any(|op|
            matches!(op, LoopSuperop::LinearMemoryV128Copy { .. })));
    }
}

#[test]
fn simd_copy_memarg_displacements_preserve_nonwrapping_addresses_and_trap_prefixes() {
    for (src, dst, len, load_offset, store_offset) in [
        (3, 1027, 48, 5, 9), (3, 1027, 17, 127, 129),
        (0, 0, 64, 0, 8), (0, 0, 64, 8, 0),
        (0, 65504, 48, 3, 16), (65504, 0, 48, 16, 3),
        (-16, 1024, 32, 32, 0), (0, -16, 32, 0, 32),
        (-16, -16, 0, 32, 32),
    ] {
        let mut expected = None;
        for (encoding, shared) in [
            (Encoding::Marked, true), (Encoding::Marked, false),
            (Encoding::Explicit, true), (Encoding::Explicit, false),
        ] {
            let mut vm = VM::new();
            vm.memory = crate::shared_memory::SharedMemory::from_vec(
                (0..65536).map(|i| (i % 251) as u8).collect());
            let _alias = shared.then(|| vm.memory.clone());
            let result = vm.run(vec![copy_loop_chunk_with_offsets(
                src, dst, len, encoding, load_offset, store_offset)]);
            let bytes = ((len as u64 + 15) / 16) * 16;
            let expected_trap = len != 0 && (
                src as u32 as u64 + u64::from(load_offset) + bytes > 65536
                || dst as u32 as u64 + u64::from(store_offset) + bytes > 65536);
            assert_eq!(result.is_err(), expected_trap,
                "nonwrapping bounds src={src} dst={dst} len={len} load={load_offset} store={store_offset}");
            let outcome = (result.is_err(), result.ok(), vm.memory.with_buffer(|b| b.to_vec()));
            if let Some(expected) = &expected { assert_eq!(&outcome, expected); }
            else { expected = Some(outcome); }
        }
    }
}

#[test]
fn cached_nonzero_simd_offsets_are_retained_in_loop_metadata() {
    let mut vm = VM::new();
    vm.memory = crate::shared_memory::SharedMemory::from_vec(vec![0x5a; 65536]);
    assert_eq!(vm.run(vec![copy_loop_chunk_with_offsets(
        3, 1027, 48, Encoding::Marked, 5, 129)]).unwrap(), Value::V128([0x5a; 16]));
    assert!(vm.loop_superops[0].iter().any(|op| matches!(op,
        LoopSuperop::LinearMemoryV128Copy { load_mem_offset: 5, store_mem_offset: 129, .. })));
}

#[test]
#[ignore = "paired displaced SIMD copies; scalar control includes alias locking"]
fn displaced_simd_copy_loop_execution_benchmark() {
    let mut samples = [Vec::new(), Vec::new()];
    for round in 0..23 {
        for step in 0..2 {
            let index = (round + step) % 2;
            let chunk = copy_loop_chunk_with_offsets(
                3, 32639, 32752, Encoding::Marked, 5, 129);
            let mut vm = VM::new();
            let seed: Vec<u8> = (0..65536).map(|i| (i % 251) as u8).collect();
            let expected = seed[8..32760].to_vec();
            let last: [u8; 16] = expected[32736..].try_into().unwrap();
            vm.memory = crate::shared_memory::SharedMemory::from_vec(seed);
            let _alias = (index == 1).then(|| vm.memory.clone());
            let clock = Instant::now();
            assert_eq!(vm.run(vec![chunk]).unwrap(), Value::V128(last));
            let elapsed = clock.elapsed().as_secs_f64();
            vm.memory.with_buffer(|bytes| assert_eq!(&bytes[32768..65520], expected.as_slice()));
            if round >= 2 { samples[index].push(elapsed); }
        }
    }
    for (index, values) in samples.iter_mut().enumerate() {
        values.sort_by(f64::total_cmp);
        eprintln!("displaced-simd-copy path={} vectors=2047 p50_ms={:.3} p95_ms={:.3}",
            if index == 0 { "exclusive-fused" } else { "aliased-scalar" },
            values[10] * 1000.0, values[19] * 1000.0);
    }
}

fn simd_fill_loop_chunk(dst: i32, len: i32, pattern: [u8; 16], encoding: Encoding) -> Chunk {
    let mut chunk = Chunk::new("simd-fill-loop");
    chunk.local_count = 4;
    chunk.scratch_high_water = 4;
    chunk.result_arity = 1;
    for (slot, value) in [(0, dst), (1, len), (2, 0)] {
        chunk.emit_i32_const(value, 1);
        copy_loop_local(&mut chunk, Op::LOCAL_SET, slot);
    }
    chunk.emit_op(Op::V128_CONST, 1);
    chunk.code.extend_from_slice(&pattern);
    chunk.lines.resize(chunk.code.len(), 1);
    copy_loop_local(&mut chunk, Op::LOCAL_SET, 3);
    chunk.emit_op_u8_u8(Op::BLOCK, 0, 0, 1);
    chunk.emit_op_u8_u8(Op::LOOP, 0, 0, 1);
    copy_loop_local(&mut chunk, Op::LOCAL_GET, 2);
    copy_loop_local(&mut chunk, Op::LOCAL_GET, 1);
    chunk.emit_op(Op::I32_GE_U, 1);
    chunk.emit_op(Op::BR_IF, 1);
    leb(&mut chunk, 1);
    copy_loop_local(&mut chunk, Op::LOCAL_GET, 0);
    copy_loop_local(&mut chunk, Op::LOCAL_GET, 2);
    chunk.emit_op(Op::I32_ADD, 1);
    copy_loop_local(&mut chunk, Op::LOCAL_GET, 3);
    chunk.emit_op(Op::V128_STORE, 1);
    memarg(&mut chunk, encoding, 0);
    copy_loop_local(&mut chunk, Op::LOCAL_GET, 2);
    chunk.emit_i32_const(16, 1);
    chunk.emit_op(Op::I32_ADD, 1);
    copy_loop_local(&mut chunk, Op::LOCAL_SET, 2);
    chunk.emit_op(Op::BR, 1);
    leb(&mut chunk, 0);
    chunk.emit_op(Op::END, 1);
    chunk.emit_op(Op::END, 1);
    copy_loop_local(&mut chunk, Op::LOCAL_GET, 2);
    chunk.emit_op(Op::RETURN, 1);
    chunk
}

#[test]
fn simd_repeated_stores_preserve_patterns_counters_and_trap_prefixes() {
    for pattern in [[0u8; 16], [0xff; 16], std::array::from_fn(|i| (i * 17 + 3) as u8)] {
        for (dst, len) in [(3, 48), (3, 17), (3, 0), (65504, 48), (-16, 32)] {
            let mut reference = None;
            for (encoding, shared) in [
                (Encoding::Marked, true), (Encoding::Marked, false),
                (Encoding::Omitted, false), (Encoding::Explicit, false),
            ] {
                let mut vm = VM::new();
                vm.memory = crate::shared_memory::SharedMemory::from_vec(vec![0x7b; 65536]);
                let _alias = shared.then(|| vm.memory.clone());
                let result = vm.run(vec![simd_fill_loop_chunk(dst, len, pattern, encoding)]);
                let bytes = ((len as u64 + 15) / 16) * 16;
                let trap = len != 0 && dst as u32 as u64 + bytes > 65536;
                assert_eq!(result.is_err(), trap);
                if !trap { assert_eq!(result.as_ref().unwrap(), &Value::I32(bytes as i32)); }
                let outcome = (result.is_err(), result.ok(), vm.memory.with_buffer(|b| b.to_vec()));
                if let Some(reference) = &reference { assert_eq!(&outcome, reference); }
                else { reference = Some(outcome); }
            }
        }
    }
}

#[test]
fn marked_repeated_stores_reuse_the_typed_fill_superop() {
    let mut vm = VM::new();
    vm.memory = crate::shared_memory::SharedMemory::from_vec(vec![0u8; 65536]);
    assert_eq!(vm.run(vec![simd_fill_loop_chunk(3, 48, [0x5a; 16], Encoding::Marked)]).unwrap(), Value::I32(48));
    assert!(vm.loop_superops[0].iter().any(|op| matches!(op,
        LoopSuperop::LinearMemoryTypedFill {
            byte_width: 16, value_kind: LinearMemoryTypedFillValue::V128, .. })));
}

#[test]
#[ignore = "paired SIMD repeated stores; scalar control includes alias locking"]
fn simd_repeated_store_execution_benchmark() {
    for uniform in [false, true] {
        let pattern = if uniform { [0x5a; 16] } else { std::array::from_fn(|i| (i * 17 + 3) as u8) };
        let mut samples = [Vec::new(), Vec::new()];
        for round in 0..23 {
            for step in 0..2 {
                let index = (round + step) % 2;
                let chunk = simd_fill_loop_chunk(3, 32768, pattern, Encoding::Marked);
                let mut vm = VM::new();
                vm.memory = crate::shared_memory::SharedMemory::from_vec(vec![0x7b; 65536]);
                let _alias = (index == 1).then(|| vm.memory.clone());
                let clock = Instant::now();
                assert_eq!(vm.run(vec![chunk]).unwrap(), Value::I32(32768));
                let elapsed = clock.elapsed().as_secs_f64();
                vm.memory.with_buffer(|bytes| {
                    assert!(bytes[3..32771].iter().enumerate().all(|(i, &byte)| byte == pattern[i % 16]));
                    assert!(bytes[..3].iter().chain(&bytes[32771..]).all(|&byte| byte == 0x7b));
                });
                if round >= 2 { samples[index].push(elapsed); }
            }
        }
        for (index, values) in samples.iter_mut().enumerate() {
            values.sort_by(f64::total_cmp);
            eprintln!("simd-repeated-store uniform={uniform} path={} vectors=2048 p50_ms={:.3} p95_ms={:.3}",
                if index == 0 { "exclusive-fused" } else { "aliased-scalar" },
                values[10] * 1000.0, values[19] * 1000.0);
        }
    }
}

// Match the nested-loop shape of the portable WASM overlap workload without
// plugin registration or module loading obscuring the execution comparison.
fn nested_overlap_copy(encoding: Encoding, rounds: i32) -> Chunk {
    let mut inner = copy_loop_chunk(0, 1, 32768, encoding);
    // Remove the final local.get(last)/return, retaining both inner ENDs.
    inner.code.truncate(inner.code.len() - 10);
    let mut chunk = Chunk::new("nested-overlap-copy");
    chunk.local_count = 7;
    chunk.scratch_high_water = 7;
    chunk.result_arity = 1;
    chunk.emit_i32_const(0, 1);
    copy_loop_local(&mut chunk, Op::LOCAL_SET, 5);
    chunk.emit_op_u8_u8(Op::BLOCK, 0, 0, 1);
    chunk.emit_op_u8_u8(Op::LOOP, 0, 0, 1);
    copy_loop_local(&mut chunk, Op::LOCAL_GET, 5);
    chunk.emit_i32_const(rounds, 1);
    chunk.emit_op(Op::I32_GE_U, 1);
    chunk.emit_op(Op::BR_IF, 1);
    leb(&mut chunk, 1);
    chunk.code.extend_from_slice(&inner.code);
    copy_loop_local(&mut chunk, Op::LOCAL_GET, 5);
    chunk.emit_i32_const(1, 1);
    chunk.emit_op(Op::I32_ADD, 1);
    copy_loop_local(&mut chunk, Op::LOCAL_SET, 5);
    chunk.emit_op(Op::BR, 1);
    leb(&mut chunk, 0);
    chunk.emit_op(Op::END, 1);
    chunk.emit_op(Op::END, 1);
    copy_loop_local(&mut chunk, Op::LOCAL_GET, 5);
    chunk.emit_op(Op::RETURN, 1);
    chunk.lines.resize(chunk.code.len(), 1);
    chunk
}

fn execute_nested_overlap(encoding: Encoding, shared: bool, rounds: i32)
    -> (f64, Vec<u8>)
{
    let mut vm = VM::new();
    vm.memory = crate::shared_memory::SharedMemory::from_vec(
        (0usize..65536).map(|i| (i.wrapping_mul(37) + 11) as u8).collect());
    let _alias = shared.then(|| vm.memory.clone());
    let clock = Instant::now();
    assert_eq!(vm.run(vec![nested_overlap_copy(encoding, rounds)]).unwrap(),
        Value::I32(rounds));
    let elapsed = clock.elapsed().as_secs_f64();
    let recognized = vm.loop_superops[0].iter().any(|op|
        matches!(op, LoopSuperop::LinearMemoryV128Copy { .. }));
    assert_eq!(recognized, !matches!(encoding, Encoding::Explicit));
    (elapsed, vm.memory.with_buffer(|bytes| bytes.to_vec()))
}

#[test]
fn nested_overlap_copy_preserves_scalar_results_and_recognition() {
    let mut expected: Vec<u8> = (0usize..65536)
        .map(|i| (i.wrapping_mul(37) + 11) as u8).collect();
    for _ in 0..17 {
        for offset in (0..32768).step_by(16) {
            let snapshot: [u8; 16] = expected[offset..offset + 16].try_into().unwrap();
            expected[offset + 1..offset + 17].copy_from_slice(&snapshot);
        }
    }
    for encoding in [Encoding::Marked, Encoding::Explicit] {
        for shared in [false, true] {
            assert_eq!(execute_nested_overlap(encoding, shared, 17).1, expected);
        }
    }
}

#[test]
#[ignore = "nested SIMD overlap execution timing, includes runtime decode/setup"]
fn nested_overlap_copy_execution_benchmark() {
    for shared in [false, true] {
        let mut samples = [Vec::new(), Vec::new()];
        let mut expected = None;
        for pair in 0..9 {
            for step in 0..2 {
                let index = (pair + step) % 2;
                let encoding = [Encoding::Marked, Encoding::Explicit][index];
                let (elapsed, bytes) = execute_nested_overlap(encoding, shared, 128);
                if let Some(expected) = &expected { assert_eq!(&bytes, expected); }
                else { expected = Some(bytes); }
                if pair >= 2 { samples[index].push(elapsed); }
            }
        }
        for samples in &mut samples { samples.sort_by(f64::total_cmp); }
        eprintln!("nested-overlap shared={shared} fused_ms={:.3} scalar_ms={:.3} scalar_over_fused={:.2}",
            samples[0][3] * 1000.0, samples[1][3] * 1000.0,
            samples[1][3] / samples[0][3]);
    }
}
