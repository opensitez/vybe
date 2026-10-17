// Exercise the production kernel, including its architecture-specific paths.
#[path = "../src/v128_copy.rs"]
mod kernel;

#[test]
fn bounded_randomized_copies_match_ordered_vector_execution() {
    let mut state = 0x74b2_158c_e063_9ad7u64;
    let mut next = || {
        state ^= state << 13;
        state ^= state >> 7;
        state ^= state << 17;
        state
    };
    for case in 0..4096 {
        let bytes = 16 * (1 + (next() % 128) as usize);
        let src = (next() % 128) as usize;
        let dst = if case % 2 == 0 {
            src + 1 + (next() % 32) as usize
        } else {
            (next() % 256) as usize
        };
        // Keep each fixture below 3 KiB, independent of host memory pressure.
        let mut actual: Vec<u8> = (0..bytes + 256).map(|_| next() as u8).collect();
        let mut expected = actual.clone();
        let mut last = [0u8; 16];
        for offset in (0..bytes).step_by(16) {
            last.copy_from_slice(&expected[src + offset..src + offset + 16]);
            expected[dst + offset..dst + offset + 16].copy_from_slice(&last);
        }
        assert_eq!(kernel::copy_forward_vectors(&mut actual, src, dst, bytes), Some(last));
        assert_eq!(actual, expected, "case={case} src={src} dst={dst} bytes={bytes}");
    }
}

#[test]
fn forward_overlap_exact_destination_end_matches_ordered_load_store() {
    for gap in 1..=16 {
        for src in [0usize, 1, 7, 31] {
            for bytes in [16usize, 32, 64, 112, 128, 144, 256, 512, 4096, 32768] {
                let dst = src + gap;
                let original: Vec<u8> = (0usize..dst + bytes)
                    .map(|i| (i.wrapping_mul(37) ^ (i.wrapping_mul(991) >> 3)) as u8)
                    .collect();
                let mut expected = original.clone();
                let mut actual = original;
                let mut last = [0u8; 16];
                for offset in (0..bytes).step_by(16) {
                    last.copy_from_slice(&expected[src + offset..src + offset + 16]);
                    expected[dst + offset..dst + offset + 16].copy_from_slice(&last);
                }
                assert_eq!(kernel::copy_forward_vectors(&mut actual, src, dst, bytes), Some(last));
                assert_eq!(actual, expected, "gap={gap} src={src} bytes={bytes}");
            }
        }
    }
}
