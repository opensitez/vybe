use crate::{Chunk, Value, VM};
use crate::opcode::Op;
use crate::shared_memory::SharedMemory;

fn wait_program64(op: Op, memory: u8, addr: i64, offset: u64) -> Chunk {
    let mut chunk = Chunk::new("atomic-memory64-index");
    chunk.emit_op(Op::I64_CONST, 1);
    let mut signed = addr;
    loop {
        let byte = (signed & 127) as u8;
        signed >>= 7;
        let done = (signed == 0 && byte & 64 == 0)
            || (signed == -1 && byte & 64 != 0);
        chunk.code.push(byte | if done { 0 } else { 128 });
        if done { break; }
    }
    if op == Op::MEMORY_ATOMIC_WAIT32 {
        chunk.emit_i32_const(9, 1);
    } else {
        chunk.emit_op(Op::I64_CONST, 1);
        chunk.code.push(9);
    }
    chunk.emit_op(Op::I64_CONST, 1);
    chunk.code.push(0);
    chunk.emit_op(op, 1);
    chunk.emit_leb_u32(if op == Op::MEMORY_ATOMIC_WAIT32 { 0xc2 } else { 0xc3 }, 1);
    let mut unsigned = offset;
    loop {
        let byte = (unsigned & 127) as u8;
        unsigned >>= 7;
        chunk.code.push(byte | if unsigned == 0 { 0 } else { 128 });
        if unsigned == 0 { break; }
    }
    chunk.code.push(memory);
    chunk.emit_op(Op::RETURN, 1);
    chunk.lines.resize(chunk.code.len(), 1);
    chunk
}

#[test]
fn memory64_atomic_wait_preserves_offset_and_selected_memory() {
    for op in [Op::MEMORY_ATOMIC_WAIT32, Op::MEMORY_ATOMIC_WAIT64] {
        for offset in [0u64, 8, 128, 16384] {
            for addr in [0i64, 8] {
                let at = offset as usize + addr as usize;
                let mut first = vec![0; at + 8];
                let mut second = vec![0; at + 8];
                first[at..at + 8].copy_from_slice(&7i64.to_le_bytes());
                second[at..at + 8].copy_from_slice(&9i64.to_le_bytes());
                for memory in [0u8, 1] {
                    let mut vm = VM::new();
                    vm.memory = SharedMemory::from_vec(first.clone());
                    vm.extra_memories = vec![SharedMemory::from_vec(second.clone())];
                    let result = vm.run(vec![wait_program64(op, memory, addr, offset)]).unwrap();
                    assert_eq!(result, Value::I32(if memory == 0 { 1 } else { 2 }),
                        "op={op:?} memory={memory} addr={addr} offset={offset}");
                    assert!(vm.memory.with_buffer(|b| b == &first));
                    assert!(vm.extra_memories[0].with_buffer(|b| b == &second));
                }
            }
        }
    }
}

#[test]
fn memory64_atomic_wait_traps_without_modifying_memory() {
    for op in [Op::MEMORY_ATOMIC_WAIT32, Op::MEMORY_ATOMIC_WAIT64] {
        for memory in [0u8, 1, 2] {
            for (addr, offset) in [(1, 0), (24, 0), (-1, 0), (8, u64::MAX), (0, u64::MAX)] {
                let mut vm = VM::new();
                let first = vec![7; 16];
                let second = vec![9; 16];
                vm.memory = SharedMemory::from_vec(first.clone());
                vm.extra_memories = vec![SharedMemory::from_vec(second.clone())];
                assert!(vm.run(vec![wait_program64(op, memory, addr, offset)]).unwrap_err().is_trap(),
                    "op={op:?} memory={memory} addr={addr} offset={offset}");
                assert!(vm.memory.with_buffer(|b| b == &first));
                assert!(vm.extra_memories[0].with_buffer(|b| b == &second));
            }
        }
    }
}

fn wait_program(op: Op, memory: u8, expected: i32, addr: i32) -> Chunk {
    let mut chunk = Chunk::new("atomic-memory-index");
    chunk.emit_i32_const(addr, 1);
    if op == Op::MEMORY_ATOMIC_WAIT32 {
        chunk.emit_i32_const(expected, 1);
    } else {
        chunk.emit_op(Op::I64_CONST, 1);
        // Test expectations 7 and 9 are single-byte signed LEB values.
        chunk.code.push(expected as u8);
    }
    chunk.emit_op(Op::I64_CONST, 1);
    chunk.code.push(0); // Zero timeout: never park a thread.
    chunk.emit_op(op, 1);
    chunk.code.extend_from_slice(&[
        if op == Op::MEMORY_ATOMIC_WAIT32 { 0x42 } else { 0x43 },
        0, memory,
    ]);
    chunk.emit_op(Op::RETURN, 1);
    chunk.lines.resize(chunk.code.len(), 1);
    chunk
}

#[test]
fn atomic_wait_reads_the_selected_memory_not_memory_zero() {
    for op in [Op::MEMORY_ATOMIC_WAIT32, Op::MEMORY_ATOMIC_WAIT64] {
        for (memory, expected, result) in [(0, 7, 2), (0, 9, 1), (1, 9, 2), (1, 7, 1)] {
            let mut vm = VM::new();
            vm.memory = SharedMemory::from_vec(7i64.to_le_bytes().to_vec());
            vm.extra_memories = vec![SharedMemory::from_vec(9i64.to_le_bytes().to_vec())];
            assert_eq!(vm.run(vec![wait_program(op, memory, expected, 0)]).unwrap(), Value::I32(result),
                "op={op:?} memory={memory} expected={expected}");
        }
    }
}

#[test]
fn selected_atomic_wait_memory_is_borrowed_without_cloning() {
    let mut vm = VM::new();
    vm.extra_memories = vec![SharedMemory::from_vec(vec![0; 8])];
    assert!(std::ptr::eq(vm.atomic_wait_memory(0).unwrap(), &vm.memory));
    assert!(std::ptr::eq(vm.atomic_wait_memory(1).unwrap(), &vm.extra_memories[0]));
    assert!(matches!(vm.atomic_wait_memory(2), Err(error) if error.is_trap()));
}

#[test]
fn atomic_wait_preserves_alignment_and_bounds_traps_for_all_memories() {
    for op in [Op::MEMORY_ATOMIC_WAIT32, Op::MEMORY_ATOMIC_WAIT64] {
        for memory in [0, 1, 2] {
            for addr in [1, 8] {
                let mut vm = VM::new();
                vm.memory = SharedMemory::from_vec(7i64.to_le_bytes().to_vec());
                vm.extra_memories = vec![SharedMemory::from_vec(9i64.to_le_bytes().to_vec())];
                assert!(vm.run(vec![wait_program(op, memory, 9, addr)]).unwrap_err().is_trap());
                assert_eq!(vm.memory.with_buffer(|b| b.to_vec()), 7i64.to_le_bytes());
                assert_eq!(vm.extra_memories[0].with_buffer(|b| b.to_vec()), 9i64.to_le_bytes());
            }
        }
    }
}

#[test]
fn atomic_notify_with_no_waiters_preserves_selected_memory() {
    for memory in [0u8, 1] {
        let mut vm = VM::new();
        vm.memory = SharedMemory::from_vec(7i64.to_le_bytes().to_vec());
        vm.extra_memories = vec![SharedMemory::from_vec(9i64.to_le_bytes().to_vec())];
        let mut chunk = Chunk::new("notify-memory-index");
        chunk.emit_i32_const(0, 1);
        chunk.emit_i32_const(0, 1);
        chunk.emit_op(Op::MEMORY_ATOMIC_NOTIFY, 1);
        chunk.code.extend_from_slice(&[0x42, 0, memory]);
        chunk.emit_op(Op::RETURN, 1);
        chunk.lines.resize(chunk.code.len(), 1);
        assert_eq!(vm.run(vec![chunk]).unwrap(), Value::I32(0));
        assert_eq!(vm.memory.with_buffer(|b| b.to_vec()), 7i64.to_le_bytes());
        assert_eq!(vm.extra_memories[0].with_buffer(|b| b.to_vec()), 9i64.to_le_bytes());
    }
}
