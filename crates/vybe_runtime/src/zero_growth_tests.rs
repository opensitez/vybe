use super::*;

#[test]
fn unchanged_growth_preserves_bytes_and_aliases() {
    for len in [0, 65536, 131072] {
        for shared in [false, true] {
            let seed = vec![0x71; len];
            let mut memory = SharedMemory::from_vec(seed.clone());
            let alias = shared.then(|| memory.clone());
            assert_eq!(memory.grow(0), len / 65536);
            assert_eq!(memory.grow_exclusive_or_shared(0), len / 65536);
            assert_eq!(memory.with_buffer(|bytes| bytes.to_vec()), seed);
            if let Some(alias) = alias {
                assert_eq!(alias.with_buffer(|bytes| bytes.to_vec()), seed);
            }
        }
    }
}

#[test]
fn unchanged_growth_keeps_declared_limit_and_overflow_failure_checks() {
    let seed = vec![0x72; 65536];
    let mut memory = SharedMemory::from_vec(seed.clone());
    memory.max_pages = Some(0);
    assert_eq!(memory.grow(0), usize::MAX);
    assert_eq!(memory.grow_exclusive_or_shared(0), usize::MAX);
    memory.max_pages = None;
    assert_eq!(memory.grow_exclusive_or_shared(usize::MAX), usize::MAX);
    assert_eq!(memory.with_buffer(|bytes| bytes.to_vec()), seed);
}
