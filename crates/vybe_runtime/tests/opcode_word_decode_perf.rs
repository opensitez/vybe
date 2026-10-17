//! Fixed-width internal opcode decoding: equivalence and paired measurements.
//! This is not a standard-WASM LEB decoder or a whole-VM throughput benchmark.
use std::hint::black_box;
use std::time::Instant;
use vybe_runtime::Op;

#[inline]
fn previous_decode(code: &[u8], offset: usize) -> Op {
    Op(((code[offset] as u32) << 24)
        | ((code[offset + 1] as u32) << 16)
        | ((code[offset + 2] as u32) << 8)
        | code[offset + 3] as u32)
}

#[test]
fn fixed_word_decode_preserves_all_opcode_bits_and_byte_order() {
    for bit in 0..32 {
        for word in [1u32 << bit, !(1u32 << bit)] {
            let op = Op(word);
            assert_eq!(Op::from_code_at(&op.encode(), 0), op);
        }
    }
    for group in 0..=u16::MAX {
        let op = Op::new(group, !group);
        assert_eq!(Op::from_code_at(&op.encode(), 0), op);
    }
}

#[test]
fn fixed_word_decode_matches_previous_at_unaligned_and_end_offsets() {
    let mut state = 0x91a2_b3c4u32;
    let code: Vec<_> = (0..8195).map(|_| {
        state ^= state << 13;
        state ^= state >> 17;
        state ^= state << 5;
        state as u8
    }).collect();
    for offset in 0..=code.len() - 4 {
        assert_eq!(Op::from_code_at(&code, offset), previous_decode(&code, offset));
    }
    let exact = &code[3..7];
    assert_eq!(Op::from_code_at(exact, 0), previous_decode(exact, 0));
}

#[test]
fn fixed_word_decode_still_rejects_incomplete_windows() {
    for len in 0usize..=7 {
        let code = vec![0u8; len];
        for offset in 0..=len + 1 {
            let old = std::panic::catch_unwind(|| previous_decode(&code, offset));
            let new = std::panic::catch_unwind(|| Op::from_code_at(&code, offset));
            assert_eq!(new.is_ok(), old.is_ok());
            if let (Ok(old), Ok(new)) = (old, new) { assert_eq!(new, old); }
        }
        assert!(std::panic::catch_unwind(|| Op::from_code_at(&code, usize::MAX)).is_err());
    }
}

fn measure(code: &[u8], current: bool) -> (f64, u32) {
    let start = Instant::now();
    let mut checksum = 0u32;
    for iteration in 0..20_000_000usize {
        let offset = black_box((iteration & 1023) * 4);
        let code = black_box(code);
        let op = if current { Op::from_code_at(code, offset) }
            else { previous_decode(code, offset) };
        checksum = checksum.wrapping_add(op.0);
    }
    black_box(checksum);
    (start.elapsed().as_secs_f64(), checksum)
}

#[test]
#[ignore = "paired fixed-opcode helper benchmark, not SIMD/VM/Doom throughput"]
fn fixed_word_decode_loaded_benchmark() {
    let code: Vec<_> = (0usize..4096).map(|i| i.wrapping_mul(31) as u8).collect();
    let mut ratios = [0.0f64; 21];
    let mut previous = [0.0f64; 21];
    let mut current = [0.0f64; 21];
    for pair in 0..23 {
        let (old, new) = if pair % 2 == 0 {
            (measure(&code, false), measure(&code, true))
        } else {
            let new = measure(&code, true);
            (measure(&code, false), new)
        };
        assert_eq!(old.1, new.1);
        if pair >= 2 {
            let sample = pair - 2;
            ratios[sample] = old.0 / new.0;
            previous[sample] = old.0;
            current[sample] = new.0;
        }
    }
    ratios.sort_by(f64::total_cmp);
    previous.sort_by(f64::total_cmp);
    current.sort_by(f64::total_cmp);
    eprintln!(
        "fixed word decode: paired_median={:.3} previous_p50_ms={:.3} current_p50_ms={:.3} previous_p95_ms={:.3} current_p95_ms={:.3}; ratios={ratios:?}",
        ratios[10], previous[10] * 1000.0, current[10] * 1000.0,
        previous[19] * 1000.0, current[19] * 1000.0,
    );
}
