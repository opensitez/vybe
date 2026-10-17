use super::*;

fn leb(code: &mut Vec<u8>, mut value: u64) {
    loop {
        let byte = (value & 0x7f) as u8;
        value >>= 7;
        code.push(byte | if value != 0 { 0x80 } else { 0 });
        if value == 0 { break; }
    }
}

fn immediate(chunk: &mut Chunk, flags: u32, offset: u64, memory: Option<u32>) -> usize {
    let start = chunk.code.len();
    leb(&mut chunk.code, flags as u64);
    leb(&mut chunk.code, offset);
    if let Some(memory) = memory { leb(&mut chunk.code, memory as u64); }
    chunk.lines.resize(chunk.code.len(), 0);
    chunk.code.len() - start
}

fn cached(op: Op, flags: u32, offset: u64, memory: Option<u32>, padding: bool) -> (Option<(u64, usize)>, usize) {
    let mut chunk = Chunk::new("<simd-memarg-cache>");
    if padding { chunk.emit_i32_const(0, 0); }
    let start = chunk.code.len();
    chunk.emit_op(op, 0);
    let size = immediate(&mut chunk, flags, offset, memory);
    chunk.emit_op(Op::NOP, 0);
    let (ops, operands, wide) = VM::build_decoded_caches_for_chunk(&chunk, 0);
    let mut vm = VM::new();
    vm.decoded_ops = vec![ops];
    vm.decoded_operands = vec![operands];
    vm.decoded_operands64 = vec![wide];
    (vm.cached_optional_memory_arg(0, start + 4), size)
}

#[test]
fn regular_simd_memory_immediates_reuse_compact_cache_at_unaligned_starts() {
    for sub in 0..256 {
        let Some(op) = Op::decode(0xfd, sub) else { continue; };
        if op.operand_format() != crate::opcode::OperandFormat::SimdMemArg { continue; }
        for padding in [false, true] {
            for offset in [0, 1, 127, 128, 16383, DECODED_MEMARG_OFFSET_MASK as u64] {
                let (actual, size) = cached(op, 0x84, offset, None, padding);
                assert_eq!(actual, Some((offset, size)), "{op:?} offset={offset}");
            }
        }
    }
}

#[test]
fn wide_explicit_and_nonmarker_immediates_keep_legacy_reader() {
    for op in [Op::V128_LOAD, Op::V128_STORE, Op::I32_LOAD] {
        for (flags, offset, memory) in [
            (0x184, 0x1_0000_0000, None),
            (0xc4, 128, Some(1)),
            (0xc4, 128, Some(u32::MAX)),
            (0x84, DECODED_MEMARG_OFFSET_MASK as u64 + 1, None),
            (0x40, 0, None),
        ] {
            assert_eq!(cached(op, flags, offset, memory, true).0, None);
        }
    }
}

#[test]
fn cached_marked_simd_load_store_advances_cursor_and_preserves_bytes() {
    for explicit_memory in [false, true] {
        let seed: Vec<u8> = (0..64).collect();
        let mut expected = seed.clone();
        expected[33..49].copy_from_slice(&seed[2..18]);
        let mut vm = VM::new();
        vm.memory = crate::shared_memory::SharedMemory::from_vec(seed.clone());
        vm.extra_memories = vec![crate::shared_memory::SharedMemory::from_vec(seed)];
        let mut chunk = Chunk::new("<simd-memarg-execution>");
        chunk.emit_i32_const(29, 0);
        chunk.emit_i32_const(0, 0);
        chunk.emit_op(Op::V128_LOAD, 0);
        immediate(&mut chunk, 0x84, 2, None);
        chunk.emit_op(Op::V128_STORE, 0);
        immediate(&mut chunk, if explicit_memory { 0xc4 } else { 0x84 }, 4,
            explicit_memory.then_some(1));
        chunk.emit_i32_const(121, 0);
        chunk.emit_op(Op::RETURN, 0);
        assert_eq!(vm.run(vec![chunk]).unwrap(), Value::I32(121));
        if explicit_memory { assert_eq!(vm.extra_memories[0].with_buffer(|bytes| bytes.to_vec()), expected); }
        else { assert_eq!(vm.memory.with_buffer(|bytes| bytes.to_vec()), expected); }
    }
}
