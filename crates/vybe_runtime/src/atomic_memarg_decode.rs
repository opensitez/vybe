//! Atomic memory-immediate decoding. Execution bounds/alignment checks remain
//! in the shared atomic access path; this only decodes bytes and advances ip.
use super::{read_leb_u32, read_leb_u64};

#[inline]
pub(super) fn decode(code: &[u8], ip: &mut usize) -> (u64, usize, bool) {
    if let Some(bytes) = code.get(*ip..).and_then(|bytes| bytes.get(..2)) {
        if bytes[0] < 0x80 {
            let flags = bytes[0];
            *ip += 1;
            let offset = if bytes[1] < 0x80 {
                *ip += 1;
                u64::from(bytes[1])
            } else { read_leb_u32(code, ip) as u64 };
            let memory = if flags & 0x40 != 0 { read_leb_u32(code, ip) as usize } else { 0 };
            return (offset, memory, false);
        }
    }
    // Canonical Memory64 flags occupy two LEB bytes. The common zero/small
    // offset can be decoded without entering the wide general decoder.
    // Do not advance the cursor until the whole prefix has been recognized:
    // incomplete/noncanonical inputs retain the general decoder's behavior.
    if let Some(bytes) = code.get(*ip..).and_then(|bytes| bytes.get(..3)) {
        if bytes[0] >= 0x80 && bytes[1] == 1 {
            let flags = bytes[0];
            if bytes[2] >= 0x80 {
                *ip += 2;
                return wide_offset(flags, code, ip);
            }
            let offset = u64::from(bytes[2]);
            *ip += 3;
            let memory = if flags & 0x40 != 0 {
                read_leb_u32(code, ip) as usize
            } else {
                0
            };
            return (offset, memory, true);
        }
    }
    general(code, ip)
}

// Canonical Memory64 flags have already been consumed. Keep the wide offset
// loop cold without decoding those flags a second time.
#[inline(never)]
fn wide_offset(flags: u8, code: &[u8], ip: &mut usize) -> (u64, usize, bool) {
    let offset = read_leb_u64(code, ip);
    let memory = if flags & 0x40 != 0 { read_leb_u32(code, ip) as usize } else { 0 };
    (offset, memory, true)
}

// Keep wide/noncanonical flags and memory64 off the common call frame.
#[inline(never)]
fn general(code: &[u8], ip: &mut usize) -> (u64, usize, bool) {
    let align = read_leb_u32(code, ip);
    let memory64 = align & 0x80 != 0;
    let offset = if memory64 { read_leb_u64(code, ip) } else { read_leb_u32(code, ip) as u64 };
    let memory = if align & 0x40 != 0 { read_leb_u32(code, ip) as usize } else { 0 };
    (offset, memory, memory64)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[inline]
    fn short_index_candidate(code: &[u8], ip: &mut usize) -> (u64, usize, bool) {
        if let Some(bytes) = code.get(*ip..).and_then(|bytes| bytes.get(..2)) {
            if bytes[0] < 0x80 {
                let flags = bytes[0];
                *ip += 1;
                let offset = if bytes[1] < 0x80 {
                    *ip += 1;
                    u64::from(bytes[1])
                } else { read_leb_u32(code, ip) as u64 };
                let memory = if flags & 0x40 != 0 {
                    if let Some(&byte) = code.get(*ip).filter(|byte| **byte < 0x80) {
                        *ip += 1;
                        usize::from(byte)
                    } else { read_leb_u32(code, ip) as usize }
                } else { 0 };
                return (offset, memory, false);
            }
        }
        decode(code, ip)
    }

    #[test]
    fn short_index_candidate_matches_existing_values_and_truncated_cursors() {
        use std::panic::{catch_unwind, AssertUnwindSafe};
        for flags in [2, 0x42, 0x82, 0xc2] {
            for offset in [0, 127, 128, 16384] {
                for memory in [0, 1, 127, 128, u32::MAX] {
                    let bytes = encoded(flags, offset, memory);
                    for end in 0..=bytes.len() {
                        let mut before = 0;
                        let mut after = 0;
                        let old = catch_unwind(AssertUnwindSafe(|| decode(&bytes[..end], &mut before)));
                        let new = catch_unwind(AssertUnwindSafe(|| short_index_candidate(&bytes[..end], &mut after)));
                        assert_eq!(old.is_err(), new.is_err());
                        if let (Ok(old), Ok(new)) = (old, new) { assert_eq!(old, new); }
                        assert_eq!(before, after);
                    }
                }
            }
        }
    }

    #[test]
    #[ignore = "paired explicit memory-index decode candidate; no production change"]
    fn short_index_candidate_benchmark() {
        use std::{hint::black_box, time::Instant};
        let measure = |kernel: fn(&[u8], &mut usize) -> (u64, usize, bool), bytes: &[u8]| {
            let start = Instant::now();
            for _ in 0..262144 {
                let mut ip = 0;
                black_box(kernel(black_box(bytes), black_box(&mut ip)));
                black_box(ip);
            }
            start.elapsed().as_secs_f64()
        };
        for bytes in [encoded(2, 0, 0), encoded(0x42, 0, 0), encoded(0x42, 3, 1),
            encoded(0x42, 127, 127), encoded(0x42, 3, 128), encoded(0x42, 16384, u32::MAX),
            encoded(0x82, 0, 0), encoded(0xc2, 3, 1)] {
            let mut ratios = Vec::new();
            for pair in 0..21 {
                let (old, new) = if pair % 2 == 0 {
                    (measure(decode, &bytes), measure(short_index_candidate, &bytes))
                } else {
                    let new = measure(short_index_candidate, &bytes);
                    (measure(decode, &bytes), new)
                };
                ratios.push(new / old);
            }
            ratios.sort_by(f64::total_cmp);
            println!("short-index bytes={bytes:?} candidate_over_current={:.4}", ratios[10]);
        }
    }

    fn leb(out: &mut Vec<u8>, mut value: u64) {
        loop {
            let byte = (value & 127) as u8;
            value >>= 7;
            out.push(byte | if value == 0 { 0 } else { 128 });
            if value == 0 { break; }
        }
    }

    fn encoded(flags: u32, offset: u64, memory: u32) -> Vec<u8> {
        let mut code = Vec::new();
        leb(&mut code, flags as u64);
        leb(&mut code, offset);
        if flags & 0x40 != 0 { leb(&mut code, memory as u64); }
        code
    }

    #[test]
    fn fast_atomic_arguments_preserve_values_and_cursor_positions() {
        for flags in 0..=255u32 {
            for offset in [0, 1, 127, 128, 16383, 16384, u32::MAX as u64, u64::MAX] {
                if flags & 0x80 == 0 && offset > u32::MAX as u64 { continue; }
                for memory in [0, 1, 127, 128, u32::MAX] {
                    let bytes = encoded(flags, offset, memory);
                    for prefix in 0..=3 {
                        let mut code = vec![0x55; prefix];
                        code.extend_from_slice(&bytes);
                        code.extend([0x7f; 4]);
                        let mut old = prefix;
                        let mut new = prefix;
                        assert_eq!(decode(&code, &mut new), general(&code, &mut old));
                        assert_eq!(new, old);
                    }
                }
            }
        }
    }

    #[test]
    fn incomplete_and_noncanonical_atomic_arguments_preserve_legacy_behavior() {
        use std::panic::{AssertUnwindSafe, catch_unwind};
        for bytes in [
            vec![], vec![2], vec![0x42, 0], vec![0x82, 0, 0],
            vec![0x82, 1, 0], vec![0xc2, 1, 0, 1],
            vec![0xc2, 1, 127, 0x80, 1],
            vec![2, 0x80, 0], vec![0x42, 0, 0x81, 0],
            vec![0x82, 1, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 1],
        ] {
            for end in 0..=bytes.len() {
                let mut old = 0;
                let mut new = 0;
                let previous = catch_unwind(AssertUnwindSafe(|| general(&bytes[..end], &mut old)));
                let current = catch_unwind(AssertUnwindSafe(|| decode(&bytes[..end], &mut new)));
                assert_eq!(current.is_err(), previous.is_err());
                if let (Ok(current), Ok(previous)) = (current, previous) { assert_eq!(current, previous); }
                assert_eq!(new, old);
            }
        }
    }

    #[test]
    #[ignore = "paired atomic memory argument decoder timing"]
    fn atomic_argument_decode_benchmark() {
        use std::{hint::black_box, time::Instant};
        // Retain the pre-Memory64 fast decoder as a test-only baseline. A
        // general-decoder comparison would attribute existing Memory32 wins
        // to this experiment and miss regressions relative to the old path.
        #[inline]
        fn previous_decode(code: &[u8], ip: &mut usize) -> (u64, usize, bool) {
            if let Some(bytes) = code.get(*ip..).and_then(|bytes| bytes.get(..2)) {
                if bytes[0] < 0x80 {
                    let flags = bytes[0];
                    *ip += 1;
                    let offset = if bytes[1] < 0x80 {
                        *ip += 1;
                        u64::from(bytes[1])
                    } else { read_leb_u32(code, ip) as u64 };
                    let memory = if flags & 0x40 != 0 { read_leb_u32(code, ip) as usize } else { 0 };
                    return (offset, memory, false);
                }
            }
            general(code, ip)
        }
        for bytes in [encoded(2, 0, 0), encoded(3, 127, 0), encoded(2, 16384, 0),
            encoded(0x42, 3, 1), encoded(0x82, 0, 0),
            encoded(0xc2, 3, 1), encoded(0x82, u64::MAX, 0)] {
            let measure = |kernel: fn(&[u8], &mut usize) -> (u64, usize, bool)| {
                let start = Instant::now();
                for _ in 0..65536 {
                    let mut ip = 0;
                    black_box(kernel(black_box(&bytes), black_box(&mut ip)));
                    black_box(ip);
                }
                start.elapsed().as_secs_f64() * 1e9 / 65536.0
            };
            let mut old = Vec::new();
            let mut new = Vec::new();
            for trial in 0..31 {
                if trial % 2 == 0 { old.push(measure(previous_decode)); new.push(measure(decode)); }
                else { new.push(measure(decode)); old.push(measure(previous_decode)); }
            }
            old.sort_by(f64::total_cmp);
            new.sort_by(f64::total_cmp);
            println!("atomic-memarg bytes={bytes:?} previous_ns={:.3} fast_ns={:.3} ratio={:.2}",
                old[15], new[15], old[15] / new[15]);
        }
    }
}
