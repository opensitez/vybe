//! Byte kernel for a forward loop of full-width vector loads and stores.

#[cfg(all(test, target_arch = "x86_64", target_feature = "sse2"))]
#[path = "v128_sse2_perf_tests.rs"]
mod sse2_perf_tests;

#[cfg(test)]
mod tight_boundary_tests {
    use super::copy_forward_vectors;

    #[test]
    fn native_carry_matches_snapshots_at_exact_unaligned_memory_ends() {
        for src in [0usize, 1, 3, 7, 15, 17] {
            for gap in 1..16 {
                for bytes in [112usize, 128, 144, 240, 256, 272, 496, 512, 528, 1040, 4096] {
                    let dst = src + gap;
                    // No trailing padding: the final destination vector ends
                    // exactly at the guest-visible memory boundary.
                    let initial: Vec<u8> = (0..dst + bytes)
                        .map(|i| (i as u8).wrapping_mul(37).wrapping_add(11)).collect();
                    let mut expected = initial.clone();
                    let mut actual = initial.clone();
                    let mut last = [0; 16];
                    for offset in (0..bytes).step_by(16) {
                        let snapshot = expected[src + offset..src + offset + 16].to_vec();
                        expected[dst + offset..dst + offset + 16].copy_from_slice(&snapshot);
                        last.copy_from_slice(&snapshot);
                    }
                    assert_eq!(copy_forward_vectors(&mut actual, src, dst, bytes), Some(last));
                    assert_eq!(actual, expected, "src={src} gap={gap} bytes={bytes}");
                    for (invalid_src, invalid_dst, invalid_bytes) in [
                        (src, dst + 1, bytes),
                        (src, dst, bytes + 16),
                        (initial.len() - 15, dst, 16),
                    ] {
                        let mut rejected = initial.clone();
                        assert_eq!(copy_forward_vectors(&mut rejected, invalid_src, invalid_dst, invalid_bytes), None);
                        assert_eq!(rejected, initial);
                    }
                }
            }
        }
    }
}

#[inline(always)]
pub(super) fn copy_forward_vectors(
    memory: &mut [u8],
    src: usize,
    dst: usize,
    bytes: usize,
) -> Option<[u8; 16]> {
    // The recognized loop always performs at least one complete vector.
    // Reject before mutation so its caller can retain ordinary trap order.
    if bytes == 0 || bytes % 16 != 0 {
        return None;
    }
    let src_end = src.checked_add(bytes)?;
    let dst_end = dst.checked_add(bytes)?;
    if src_end > memory.len() || dst_end > memory.len() {
        return None;
    }
    let mut last = [0u8; 16];
    if dst > src && dst < src_end {
        let gap = dst - src;
        #[cfg(debug_assertions)]
        if gap >= 16 && bytes >= 128 {
            // Debug vector/period loops incur substantial per-iteration
            // overhead. For disjoint load/store pairs, the original gap
            // repeats: initialize it once and double the written prefix.
            // Release builds retain their measured short-copy kernels.
            memory.copy_within(src..dst, dst);
            let mut written = gap;
            while written < bytes {
                let count = written.min(bytes - written);
                memory.copy_within(dst..dst + count, dst + written);
                written += count;
            }
            last.copy_from_slice(&memory[dst_end - 16..dst_end]);
            return Some(last);
        }
        if gap == 16 && (128..4096).contains(&bytes) {
            // Each forward load observes the vector just stored by the
            // preceding iteration. The first snapshot therefore repeats.
            last.copy_from_slice(&memory[src..src + 16]);
            for vector in memory[dst..dst_end].chunks_exact_mut(16) {
                vector.copy_from_slice(&last);
            }
            return Some(last);
        }
        if (256..4096).contains(&bytes) {
            #[inline]
            fn fill_period<const GAP: usize>(
                memory: &mut [u8], src: usize, dst: usize, bytes: usize,
            ) -> [u8; 16] {
                let pattern: [u8; GAP] = memory[src..src + GAP].try_into().unwrap();
                let mut blocks = memory[dst..dst + bytes].chunks_exact_mut(GAP);
                for block in &mut blocks {
                    block.copy_from_slice(&pattern);
                }
                let remainder = blocks.into_remainder();
                remainder.copy_from_slice(&pattern[..remainder.len()]);
                let last_start = (bytes - 16) % GAP;
                pattern[last_start..last_start + 16].try_into().unwrap()
            }
            // These periods are whole vectors. Snapshot the untouched
            // source period once, including the final partial-period case.
            match gap {
                32 => return Some(fill_period::<32>(memory, src, dst, bytes)),
                64 => return Some(fill_period::<64>(memory, src, dst, bytes)),
                128 => return Some(fill_period::<128>(memory, src, dst, bytes)),
                _ => {}
            }
        }
        // Measured non-debug SSE2 carry wins for these large displacements;
        // let them reach the existing checked carry helper below. Other
        // gaps, debug builds, and non-SSE2 targets retain portable kernels.
        if bytes >= 512 && matches!(gap, 1..=8)
            && !(cfg!(all(target_arch = "x86_64", target_feature = "sse2", not(debug_assertions)))
                && matches!(gap, 1 | 2 | 8)) {
            // Carry the preceding vector's suffix directly for gaps where
            // the portable register kernel beats bulk-copy-plus-repair.
            // A fresh vector starts at the current destination: it has not
            // yet been written, and is inside the checked destination range.
            #[inline(always)]
            fn carry<const SHIFT: u32>(
                memory: &mut [u8], src: usize, dst: usize, bytes: usize,
            ) -> [u8; 16] {
                let mut previous = u128::from_le_bytes(
                    memory[src..src + 16].try_into().unwrap(),
                );
                memory[dst..dst + 16].copy_from_slice(&previous.to_le_bytes());
                for offset in (16..bytes).step_by(16) {
                    let fresh = u128::from_le_bytes(
                        memory[dst + offset..dst + offset + 16].try_into().unwrap(),
                    );
                    previous = (previous >> (SHIFT * 8))
                        | (fresh << ((16 - SHIFT) * 8));
                    memory[dst + offset..dst + offset + 16]
                        .copy_from_slice(&previous.to_le_bytes());
                }
                previous.to_le_bytes()
            }
            match gap {
                3 => return Some(carry::<13>(memory, src, dst, bytes)),
                5 => return Some(carry::<11>(memory, src, dst, bytes)),
                6 => return Some(carry::<10>(memory, src, dst, bytes)),
                7 => return Some(carry::<9>(memory, src, dst, bytes)),
                _ => {}
            }
            // Each snapshot after the first sees exactly `gap` bytes from
            // the preceding store. Copy untouched bytes in bulk, then repair
            // those prefixes in order; ordinary memmove alone is incorrect.
            #[inline]
            fn repair<const GAP: usize>(memory: &mut [u8], dst: usize, bytes: usize) {
                #[cfg(all(target_arch = "aarch64", not(debug_assertions)))]
                if GAP == 1 {
                    // Carry the untouched tail in a register instead of
                    // reloading the preceding vector for each repair. A
                    // one-byte prefix repair cannot change the tail.
                    let mut vectors = memory[dst..dst + bytes].chunks_exact_mut(16);
                    let first = vectors.next().unwrap();
                    let mut previous = first[15];
                    for vector in vectors {
                        let next = vector[15];
                        vector[0] = previous;
                        previous = next;
                    }
                    return;
                }
                for offset in (16..bytes).step_by(16) {
                    let previous: [u8; GAP] = memory[dst + offset - GAP..dst + offset]
                        .try_into()
                        .unwrap();
                    memory[dst + offset..dst + offset + GAP].copy_from_slice(&previous);
                }
            }
            memory.copy_within(src..src + bytes, dst);
            match gap {
                1 => repair::<1>(memory, dst, bytes),
                2 => repair::<2>(memory, dst, bytes),
                4 => repair::<4>(memory, dst, bytes),
                8 => repair::<8>(memory, dst, bytes),
                _ => unreachable!(),
            }
            let mut last = [0; 16];
            last.copy_from_slice(&memory[dst + bytes - 16..dst + bytes]);
            return Some(last);
        }
        if gap >= 16 && bytes >= 512 {
            // Each vector load is disjoint from its own store. Sequential
            // stores therefore repeat the original gap-sized prefix, even
            // when the gap is not a multiple of the vector width. Copy the
            // untouched prefix once, then double the initialized destination.
            memory.copy_within(src..dst, dst);
            let mut written = gap;
            while written < bytes {
                let count = written.min(bytes - written);
                memory.copy_within(dst..dst + count, dst + written);
                written += count;
            }
            last.copy_from_slice(&memory[dst_end - 16..dst_end]);
            return Some(last);
        }
        // Forward overlap may propagate earlier stores into later loads.
        // Snapshot each whole vector before its store, including when the
        // displacement is smaller than sixteen bytes.
        #[cfg(all(target_arch = "aarch64", target_feature = "neon"))]
        if bytes >= 128 && (matches!(gap, 9..=15)
            || (bytes < 512 && matches!(gap, 1..=8)
                && (bytes >= 256 || cfg!(not(debug_assertions))))) {
            // Carry the preceding vector in a register rather than reloading
            // a source vector spanning the preceding store. The fresh suffix
            // starts at the current destination, which has not been written.
            // Both complete ranges were checked above; even the ignored bytes
            // of that fresh vector remain inside the destination range.
            #[inline]
            unsafe fn carry<const SHIFT: i32>(
                memory: &mut [u8],
                src: usize,
                dst: usize,
                bytes: usize,
            ) -> [u8; 16] {
                use core::arch::aarch64::{vextq_u8, vld1q_u8, vst1q_u8};
                // Safety: caller validates both ranges, bytes is a positive
                // multiple of 16, and dst - src = 16 - SHIFT (1..=15).
                // Each fresh load and output store is wholly inside dst's
                // checked range. The mutable slice excludes concurrent access.
                let ptr = memory.as_mut_ptr();
                let mut previous = unsafe { vld1q_u8(ptr.add(src)) };
                unsafe { vst1q_u8(ptr.add(dst), previous) };
                for offset in (16..bytes).step_by(16) {
                    let fresh = unsafe { vld1q_u8(ptr.add(dst + offset)) };
                    previous = unsafe { vextq_u8::<SHIFT>(previous, fresh) };
                    unsafe { vst1q_u8(ptr.add(dst + offset), previous) };
                }
                let mut result = [0u8; 16];
                unsafe { vst1q_u8(result.as_mut_ptr(), previous) };
                result
            }
            // Safety: the enclosing branch supplies the helper's exact gap,
            // size, bounds, and compile-time NEON availability requirements.
            return Some(unsafe {
                match gap {
                    1 => carry::<15>(memory, src, dst, bytes),
                    2 => carry::<14>(memory, src, dst, bytes),
                    3 => carry::<13>(memory, src, dst, bytes),
                    4 => carry::<12>(memory, src, dst, bytes),
                    5 => carry::<11>(memory, src, dst, bytes),
                    6 => carry::<10>(memory, src, dst, bytes),
                    7 => carry::<9>(memory, src, dst, bytes),
                    8 => carry::<8>(memory, src, dst, bytes),
                    9 => carry::<7>(memory, src, dst, bytes),
                    10 => carry::<6>(memory, src, dst, bytes),
                    11 => carry::<5>(memory, src, dst, bytes),
                    12 => carry::<4>(memory, src, dst, bytes),
                    13 => carry::<3>(memory, src, dst, bytes),
                    14 => carry::<2>(memory, src, dst, bytes),
                    15 => carry::<1>(memory, src, dst, bytes),
                    _ => unreachable!(),
                }
            });
        }
        // Unoptimized intrinsic wrappers regress this loop substantially;
        // retain the scalar path in debug builds. Optimized-build paired
        // measurements and both debug/release correctness gates are retained.
        #[cfg(all(target_arch = "x86_64", target_feature = "sse2", not(debug_assertions)))]
        if bytes >= 128 && (matches!(gap, 9..=15)
            || matches!(gap, 1 | 2 | 8)
            || (bytes < 512 && matches!(gap, 1..=8))) {
            #[inline]
            unsafe fn carry<const SHIFT: i32, const GAP: i32>(
                memory: &mut [u8], src: usize, dst: usize, bytes: usize,
            ) -> [u8; 16] {
                use core::arch::x86_64::{
                    _mm_loadu_si128, _mm_storeu_si128, _mm_or_si128,
                    _mm_srli_si128, _mm_slli_si128,
                };
                // Safety: both complete ranges are validated above; every
                // fresh vector lies in the checked destination range. SSE2
                // is available under this compile-time gate; loads/stores
                // are unaligned. No concurrent access to the mutable slice.
                let ptr = memory.as_mut_ptr();
                let mut previous = unsafe { _mm_loadu_si128(ptr.add(src).cast()) };
                unsafe { _mm_storeu_si128(ptr.add(dst).cast(), previous) };
                for offset in (16..bytes).step_by(16) {
                    let fresh = unsafe { _mm_loadu_si128(ptr.add(dst + offset).cast()) };
                    previous = unsafe {
                        _mm_or_si128(_mm_srli_si128::<SHIFT>(previous), _mm_slli_si128::<GAP>(fresh))
                    };
                    unsafe { _mm_storeu_si128(ptr.add(dst + offset).cast(), previous) };
                }
                let mut result = [0u8; 16];
                unsafe { _mm_storeu_si128(result.as_mut_ptr().cast(), previous) };
                result
            }
            // The shifts concatenate the previous vector's suffix and the
            // untouched fresh prefix, exactly like each forward Wasm load.
            return Some(unsafe {
                match gap {
                    1 => carry::<15, 1>(memory, src, dst, bytes),
                    2 => carry::<14, 2>(memory, src, dst, bytes),
                    3 => carry::<13, 3>(memory, src, dst, bytes),
                    4 => carry::<12, 4>(memory, src, dst, bytes),
                    5 => carry::<11, 5>(memory, src, dst, bytes),
                    6 => carry::<10, 6>(memory, src, dst, bytes),
                    7 => carry::<9, 7>(memory, src, dst, bytes),
                    8 => carry::<8, 8>(memory, src, dst, bytes),
                    9 => carry::<7, 9>(memory, src, dst, bytes),
                    10 => carry::<6, 10>(memory, src, dst, bytes),
                    11 => carry::<5, 11>(memory, src, dst, bytes),
                    12 => carry::<4, 12>(memory, src, dst, bytes),
                    13 => carry::<3, 13>(memory, src, dst, bytes),
                    14 => carry::<2, 14>(memory, src, dst, bytes),
                    15 => carry::<1, 15>(memory, src, dst, bytes),
                    _ => unreachable!(),
                }
            });
        }
        for offset in (0..bytes).step_by(16) {
            last.copy_from_slice(&memory[src + offset..src + offset + 16]);
            memory[dst + offset..dst + offset + 16].copy_from_slice(&last);
        }
    } else {
        // Backward overlap cannot overwrite a future source vector. Its
        // forward-loop result therefore equals a snapshot/memmove copy.
        last.copy_from_slice(&memory[src_end - 16..src_end]);
        if src != dst {
            memory.copy_within(src..src_end, dst);
        }
    }
    Some(last)
}

#[cfg(test)]
mod short_copy_cutoff_tests {
    use super::copy_forward_vectors;

    #[test]
    fn short_overlap_cutoffs_preserve_ordered_snapshots() {
        for src in [0usize, 3, 17] {
            for gap in (1usize..17).chain([32, 64, 128]) {
                for bytes in [112usize, 128, 144, 240, 256, 272, 496, 512, 528, 1008, 1024, 1040, 2032, 2048, 2064, 3072] {
                    let dst = src + gap;
                    let mut expected: Vec<u8> = (0..dst + bytes + 37)
                        .map(|i| (i as u8).wrapping_mul(37).wrapping_add(11))
                        .collect();
                    let mut actual = expected.clone();
                    let mut last = [0u8; 16];
                    for offset in (0..bytes).step_by(16) {
                        last.copy_from_slice(&expected[src + offset..src + offset + 16]);
                        expected[dst + offset..dst + offset + 16].copy_from_slice(&last);
                    }
                    assert_eq!(copy_forward_vectors(&mut actual, src, dst, bytes), Some(last),
                        "src={src}, gap={gap}, bytes={bytes}");
                    assert_eq!(actual, expected, "src={src}, gap={gap}, bytes={bytes}");
                }
            }
        }
    }
}

#[cfg(test)]
mod fixed_gap_perf_tests {
    use super::copy_forward_vectors;
    use std::hint::black_box;
    use std::time::Instant;

    fn previous(memory: &mut [u8], src: usize, dst: usize, bytes: usize) -> Option<[u8; 16]> {
        if bytes == 0 || bytes % 16 != 0 {
            return None;
        }
        let src_end = src.checked_add(bytes)?;
        let dst_end = dst.checked_add(bytes)?;
        if src_end > memory.len() || dst_end > memory.len() {
            return None;
        }
        if !(dst > src && dst < src_end
            && ((3..=16).contains(&(dst - src)) || matches!(dst - src, 32 | 64 | 128))) {
            return copy_forward_vectors(memory, src, dst, bytes);
        }
        if dst - src == 16 && !(128..4096).contains(&bytes) {
            return copy_forward_vectors(memory, src, dst, bytes);
        }
        if matches!(dst - src, 32 | 64 | 128) && !(256..4096).contains(&bytes) {
            return copy_forward_vectors(memory, src, dst, bytes);
        }
        let mut last = [0u8; 16];
        for offset in (0..bytes).step_by(16) {
            last.copy_from_slice(&memory[src + offset..src + offset + 16]);
            memory[dst + offset..dst + offset + 16].copy_from_slice(&last);
        }
        Some(last)
    }

    fn sample(bytes: usize, gap: usize, optimized: bool) -> f64 {
        let mut memory: Vec<u8> = (0..bytes + gap + 37)
            .map(|i| (i as u8).wrapping_mul(37).wrapping_add(11))
            .collect();
        let rounds = (8_388_608usize / bytes).max(8);
        let start = Instant::now();
        for _ in 0..rounds {
            let memory = black_box(&mut memory[..]);
            let last = if optimized {
                copy_forward_vectors(memory, 3, 3 + gap, black_box(bytes))
            } else {
                previous(memory, 3, 3 + gap, black_box(bytes))
            };
            black_box(last.unwrap());
        }
        start.elapsed().as_secs_f64()
    }

    #[test]
    #[ignore = "paired fixed-gap SIMD copy kernel measurement"]
    fn paired_fixed_gap_copy() {
        for gap in [3usize, 4, 5, 6, 7, 8, 12, 16, 32, 64, 128] {
            for bytes in [128usize, 256, 512, 1024, 2048, 4096, 32768, 1_048_576] {
                let initial: Vec<u8> = (0..bytes + gap + 37)
                    .map(|i| (i as u8).wrapping_mul(37).wrapping_add(11))
                    .collect();
                let mut old = initial.clone();
                let mut new = initial;
                assert_eq!(previous(&mut old, 3, 3 + gap, bytes), copy_forward_vectors(&mut new, 3, 3 + gap, bytes));
                assert_eq!(old, new);
                let mut ratios = [0.0f64; 11];
                for pair in 0usize..13 {
                    let (old, new) = if pair % 2 == 0 {
                        (sample(bytes, gap, false), sample(bytes, gap, true))
                    } else {
                        let new = sample(bytes, gap, true);
                        (sample(bytes, gap, false), new)
                    };
                    if pair >= 2 {
                        ratios[pair - 2] = old / new;
                    }
                }
                ratios.sort_by(f64::total_cmp);
                eprintln!("fixed SIMD gap={gap} bytes={bytes}: previous/current median {:.3}x; pairs {ratios:?}", ratios[5]);
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn reference(memory: &mut [u8], src: usize, dst: usize, bytes: usize) -> [u8; 16] {
        let mut last = [0; 16];
        for offset in (0..bytes).step_by(16) {
            // Deliberately allocated snapshot: independent sequential-load
            // reference, not a byte-at-a-time overlapping copy.
            let loaded = memory[src + offset..src + offset + 16].to_vec();
            memory[dst + offset..dst + offset + 16].copy_from_slice(&loaded);
            last.copy_from_slice(&loaded);
        }
        last
    }

    #[test]
    fn overlap_kernel_matches_sequential_vector_load_store() {
        let initial: Vec<u8> = (0..256).map(|i| (i * 37 + 11) as u8).collect();
        for src in 0..=33 {
            for dst in 0..=33 {
                for bytes in [16, 32, 64, 128] {
                    let mut expected = initial.clone();
                    let last = reference(&mut expected, src, dst, bytes);
                    let mut actual = initial.clone();
                    assert_eq!(copy_forward_vectors(&mut actual, src, dst, bytes), Some(last));
                    assert_eq!(actual, expected, "src={src} dst={dst} bytes={bytes}");
                }
            }
        }
    }

    #[test]
    fn rejected_ranges_preserve_every_byte() {
        let initial = vec![0xa5; 256];
        for (src, dst, bytes) in [
            (0, 0, 0), (0, 0, 1), (0, 0, 17),
            (usize::MAX, 0, 16), (0, usize::MAX, 16),
            (0, 0, usize::MAX - 15), (241, 0, 16),
            (0, 241, 16), (240, 0, 32), (0, 240, 32),
        ] {
            let mut actual = initial.clone();
            assert!(copy_forward_vectors(&mut actual, src, dst, bytes).is_none());
            assert_eq!(actual, initial);
        }
        let mut actual = initial;
        assert_eq!(copy_forward_vectors(&mut actual, 240, 240, 16), Some([0xa5; 16]));
    }

    #[test]
    fn chunked_overlap_preserves_unaligned_periods_and_last_vector() {
        let initial: Vec<u8> = (0usize..8192)
            .map(|i| (i.wrapping_mul(37) + 11) as u8).collect();
        for src in [0, 3, 17] {
            for gap in (1..=65).chain([127, 255, 256, 257, 511, 1023]) {
                for bytes in [240, 256, 272, 496, 512, 528, 1024, 4080, 4096, 4112] {
                    let mut expected = initial.clone();
                    let last = reference(&mut expected, src, src + gap, bytes);
                    let mut actual = initial.clone();
                    assert_eq!(copy_forward_vectors(&mut actual, src, src + gap, bytes), Some(last));
                    assert_eq!(actual, expected, "src={src} gap={gap} bytes={bytes}");
                }
            }
        }
    }

    // Previous kernel without the chunked forward-overlap branch. Keep its
    // stack snapshot, bounds checks and nonoverlap path for a fair control.
    fn previous(memory: &mut [u8], src: usize, dst: usize, bytes: usize) -> Option<[u8; 16]> {
        if bytes == 0 || bytes % 16 != 0 { return None; }
        let src_end = src.checked_add(bytes)?;
        let dst_end = dst.checked_add(bytes)?;
        if src_end > memory.len() || dst_end > memory.len() { return None; }
        let mut last = [0; 16];
        if dst > src && dst < src_end {
            for offset in (0..bytes).step_by(16) {
                last.copy_from_slice(&memory[src + offset..src + offset + 16]);
                memory[dst + offset..dst + offset + 16].copy_from_slice(&last);
            }
        } else {
            last.copy_from_slice(&memory[src_end - 16..src_end]);
            if src != dst { memory.copy_within(src..src_end, dst); }
        }
        Some(last)
    }

    #[test]
    #[ignore = "chunked SIMD overlap kernel comparison, excludes VM dispatch"]
    fn chunked_overlap_execution_benchmark() {
        use std::{hint::black_box, time::Instant};
        for bytes in [256usize, 1024, 4096, 32768, 1048576] {
            for gap in [1usize, 8, 16, 17, 31, 256, 4097] {
                let initial: Vec<u8> = (0..bytes + gap + 3)
                    .map(|i| (i.wrapping_mul(37) + 11) as u8).collect();
                let rounds = (1048576 / bytes).max(8);
                let mut samples = [Vec::new(), Vec::new()];
                for pair in 0..9 {
                    let mut expected = None;
                    for step in 0..2 {
                        let index = (pair + step) % 2;
                        let mut memory = initial.clone();
                        let mut last = None;
                        let clock = Instant::now();
                        for _ in 0..rounds {
                            last = if index == 0 {
                                copy_forward_vectors(black_box(&mut memory), 3, 3 + gap, bytes)
                            } else {
                                previous(black_box(&mut memory), 3, 3 + gap, bytes)
                            };
                            black_box(last);
                        }
                        let elapsed = clock.elapsed().as_secs_f64();
                        if let Some((expected_memory, expected_last)) = &expected {
                            assert_eq!(&memory, expected_memory);
                            assert_eq!(&last, expected_last);
                        } else { expected = Some((memory, last)); }
                        if pair >= 2 { samples[index].push(elapsed); }
                    }
                }
                for samples in &mut samples { samples.sort_by(f64::total_cmp); }
                eprintln!("chunked-overlap bytes={bytes} gap={gap} current_ms={:.3} previous_ms={:.3} previous_over_current={:.3}",
                    samples[0][3] * 1000.0, samples[1][3] * 1000.0,
                    samples[1][3] / samples[0][3]);
            }
        }
    }
}

#[cfg(test)]
mod small_gap_repair_tests {
    use super::copy_forward_vectors;

    #[test]
    fn specialized_overlap_matches_full_vector_snapshots_at_cutoffs() {
        for src in [0usize, 3, 17] {
            for gap in 1..16 {
                for bytes in [16usize, 32, 48, 240, 256, 272, 4080, 4096, 4112, 32768, 1048576] {
                    let mut state = 0x9e3779b9u32;
                    let initial: Vec<u8> = (0..src + gap + bytes + 32).map(|_| {
                        state ^= state << 13;
                        state ^= state >> 17;
                        state ^= state << 5;
                        state as u8
                    }).collect();
                    let mut expected = initial.clone();
                    let mut actual = initial;
                    let dst = src + gap;
                    let mut last = [0; 16];
                    for offset in (0..bytes).step_by(16) {
                        last.copy_from_slice(&expected[src + offset..src + offset + 16]);
                        expected[dst + offset..dst + offset + 16].copy_from_slice(&last);
                    }
                    assert_eq!(copy_forward_vectors(&mut actual, src, dst, bytes), Some(last));
                    assert_eq!(actual, expected, "src={src} gap={gap} bytes={bytes}");
                }
            }
        }
    }

    #[test]
    fn rejected_ranges_do_not_mutate_memory() {
        for (src, dst, bytes) in [
            (1usize, 2usize, 0usize),
            (1, 2, 17),
            (0, 1, 128),
            (usize::MAX - 15, 2, 32),
            (1, usize::MAX - 15, 32),
        ] {
            let mut memory: Vec<u8> = (0..96).collect();
            let before = memory.clone();
            assert_eq!(copy_forward_vectors(&mut memory, src, dst, bytes), None);
            assert_eq!(memory, before);
        }
    }
}
