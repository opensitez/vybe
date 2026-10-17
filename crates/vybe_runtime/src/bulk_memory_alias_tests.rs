//! Different memory indices may import the same physical memory instance.
use super::*;

fn aliased_machine(vm: &mut VM, bytes: &[u8]) {
    vm.memory = SharedMemory::from_vec(bytes.to_vec());
    vm.extra_memories = vec![vm.memory.clone(), vm.memory.clone()];
}

fn assert_alias_bytes(vm: &VM, expected: &[u8]) {
    assert_eq!(vm.memory.with_buffer(|bytes| bytes.to_vec()), expected);
    for memory in &vm.extra_memories {
        assert_eq!(memory.with_buffer(|bytes| bytes.to_vec()), expected);
    }
}

#[test]
fn aliased_indices_preserve_memmove_overlap_identity_and_zero_length() {
    let mut vm = VM::new();
    let seed: Vec<u8> = (0..32).collect();
    for dst_memory in 0..3 {
        for src_memory in 0..3 {
            for (dst, src, count) in [(7, 0, 20), (0, 7, 20), (5, 5, 20), (32, 32, 0)] {
                aliased_machine(&mut vm, &seed);
                let mut expected = seed.clone();
                expected.copy_within(src..src + count, dst);
                vm.copy_memory_bytes(dst_memory, dst, src_memory, src, count).unwrap();
                assert_alias_bytes(&vm, &expected);
            }
        }
    }
}

#[test]
fn aliased_indices_reject_both_ranges_before_mutation() {
    let mut vm = VM::new();
    let seed = [0x91; 32];
    for dst_memory in 0..3 {
        for src_memory in 0..3 {
            for (dst, src, count) in [(33, 0, 0), (0, 33, 0), (31, 0, 2),
                (0, 31, 2), (usize::MAX, 0, 2), (0, usize::MAX, 2)] {
                aliased_machine(&mut vm, &seed);
                assert!(vm.copy_memory_bytes(dst_memory, dst, src_memory, src, count).is_err());
                assert_alias_bytes(&vm, &seed);
            }
        }
    }
}
