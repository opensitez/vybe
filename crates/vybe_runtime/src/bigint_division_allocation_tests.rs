use super::*;
use std::{hint::black_box, time::Instant};

fn scalar_shl(a: &[u64], bits: usize) -> Vec<u64> {
    if a.is_empty() { return Vec::new(); }
    let words = bits / 64;
    let shift = bits % 64;
    let mut out = vec![0; a.len() + words + 1];
    for (i, &word) in a.iter().enumerate() {
        out[i + words] |= word << shift;
        if shift > 0 { out[i + words + 1] |= word >> (64 - shift); }
    }
    mag_norm(out)
}

fn scalar_shr(a: &[u64], bits: usize) -> Vec<u64> {
    let words = bits / 64;
    if words >= a.len() { return Vec::new(); }
    let shift = bits % 64;
    let mut out = Vec::with_capacity(a.len() - words);
    for i in words..a.len() {
        let mut word = a[i] >> shift;
        if shift > 0 {
            if let Some(&high) = a.get(i + 1) { word |= high << (64 - shift); }
        }
        out.push(word);
    }
    mag_norm(out)
}

#[test]
fn bulk_limb_shifts_match_scalar_bits_and_normalization() {
    let mut seed = 0x9876_5432_1abc_def0;
    for len in 0..32 {
        for trailing_zeros in [false, true] {
            let mut a: Vec<u64> = (0..len).map(|_| next(&mut seed)).collect();
            if trailing_zeros && len != 0 { a[len - 1] = 0; }
            for bits in 0..(len * 64 + 129) {
                assert_eq!(mag_shl(&a, bits), scalar_shl(&a, bits), "left len={len} bits={bits}");
                assert_eq!(mag_shr(&a, bits), scalar_shr(&a, bits), "right len={len} bits={bits}");
            }
        }
    }
}

#[test]
#[ignore = "paired limb-aligned bulk-copy shift benchmark"]
fn bulk_limb_shift_benchmark() {
    for len in [2usize, 16, 1024, 64000] {
        let a = vec![u64::MAX; len];
        for (name, old, new) in [
            ("left", scalar_shl as fn(&[u64], usize) -> Vec<u64>, mag_shl as fn(&[u64], usize) -> Vec<u64>),
            ("right", scalar_shr as fn(&[u64], usize) -> Vec<u64>, mag_shr as fn(&[u64], usize) -> Vec<u64>),
        ] {
            let methods = [old, new];
            let expected = old(&a, 64);
            let mut samples = [Vec::new(), Vec::new()];
            for trial in 0..31 {
                for order in 0..2 {
                    let which = (trial + order) % 2;
                    let start = Instant::now();
                    let actual = methods[which](black_box(&a), black_box(64));
                    let elapsed = start.elapsed().as_nanos();
                    assert_eq!(actual, expected);
                    samples[which].push(elapsed);
                }
            }
            for sample in &mut samples { sample.sort_unstable(); }
            println!("limb-shift direction={name} limbs={len} scalar_ns={} bulk_ns={} speedup={:.3}",
                samples[0][15], samples[1][15], samples[0][15] as f64 / samples[1][15] as f64);
        }
    }
}

// Frozen allocation-heavy algorithm for differential checks and paired timing.
fn allocating_divrem(a: &[u64], b: &[u64]) -> (Vec<u64>, Vec<u64>) {
    if mag_cmp(a, b) == std::cmp::Ordering::Less {
        return (Vec::new(), a.to_vec());
    }
    assert!(b.len() > 1);
    let mut q = vec![0; a.len()];
    let mut r = Vec::new();
    for i in (0..mag_bit_len(a)).rev() {
        r = mag_shl(&r, 1);
        if mag_get_bit(a, i) { mag_set_bit(&mut r, 0); }
        if mag_cmp(&r, b) != std::cmp::Ordering::Less {
            r = mag_sub(&r, b);
            mag_set_bit(&mut q, i);
        }
    }
    (mag_norm(q), r)
}

fn next(seed: &mut u64) -> u64 {
    *seed ^= *seed << 13;
    *seed ^= *seed >> 7;
    *seed ^= *seed << 17;
    *seed
}

#[test]
fn multi_limb_division_matches_allocating_reference_and_reconstructs_input() {
    let mut seed = 0xc0ff_ee12_3456_789a;
    for case in 0..1024 {
        let len = 2 + case % 15;
        let mut a: Vec<u64> = (0..len).map(|_| next(&mut seed)).collect();
        let mut b: Vec<u64> = (0..(2 + case % (len - 1))).map(|_| next(&mut seed)).collect();
        if case % 4 == 0 { a.fill(u64::MAX); }
        if case % 4 == 1 { a.fill(0); a[len - 1] = 1; }
        let last = b.len() - 1;
        b[last] |= 1;
        a = mag_norm(a);
        let actual = mag_divrem(&a, &b);
        assert_eq!(actual, allocating_divrem(&a, &b), "case {case}");
        assert_eq!(mag_cmp(&actual.1, &b), std::cmp::Ordering::Less);
        assert_eq!(mag_norm(mag_add(&mag_mul(&actual.0, &b), &actual.1)), a);
    }
    for bit in 64..256 {
        let mut a = vec![0; 4];
        mag_set_bit(&mut a, bit);
        a = mag_norm(a);
        let b = vec![u64::MAX, 1];
        assert_eq!(mag_divrem(&a, &b), allocating_divrem(&a, &b));
    }
}

#[test]
#[ignore = "paired allocation-heavy versus retained-remainder division"]
fn multi_limb_division_allocation_benchmark() {
    for len in [2usize, 4, 16, 64] {
        let a = vec![u64::MAX; len];
        let mut b = vec![0; len / 2 + 1];
        b[0] = 0x1234_5678_9abc_def1;
        let last = b.len() - 1;
        b[last] = 1;
        let expected = allocating_divrem(&a, &b);
        let methods = [allocating_divrem as fn(&[u64], &[u64]) -> (Vec<u64>, Vec<u64>), mag_divrem];
        let mut samples = [Vec::new(), Vec::new()];
        for trial in 0..21 {
            for order in 0..2 {
                let which = (trial + order) % 2;
                let start = Instant::now();
                let actual = methods[which](black_box(&a), black_box(&b));
                let elapsed = start.elapsed().as_nanos();
                assert_eq!(actual, expected);
                samples[which].push(elapsed);
            }
        }
        for sample in &mut samples { sample.sort_unstable(); }
        println!("bigint-divrem limbs={len} allocating_ns={} retained_ns={} speedup={:.3}",
            samples[0][10], samples[1][10], samples[0][10] as f64 / samples[1][10] as f64);
    }
}
