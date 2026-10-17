//! Exact integer SIMD kernels. Native code is limited to baseline CPU features.
//! Other targets use the same little-endian scalar lane semantics.

#[cfg(all(target_endian = "little", target_arch = "aarch64"))]
use std::arch::aarch64::*;
#[cfg(all(target_endian = "little", target_arch = "x86_64"))]
use std::arch::x86_64::*;

#[cfg(any(test, not(all(target_endian = "little", any(target_arch = "aarch64", target_arch = "x86_64")))))]
#[inline(always)]
fn scalar<const WIDTH: usize, const OP: u8>(a: &[u8; 16], b: &[u8; 16]) -> [u8; 16] {
    let mut out = [0u8; 16];
    for start in (0..16).step_by(WIDTH) {
        let mut left = [0u8; 8];
        let mut right = [0u8; 8];
        left[..WIDTH].copy_from_slice(&a[start..start + WIDTH]);
        right[..WIDTH].copy_from_slice(&b[start..start + WIDTH]);
        let left = u64::from_le_bytes(left);
        let right = u64::from_le_bytes(right);
        let result = match OP {
            0 => left.wrapping_add(right),
            1 => left.wrapping_sub(right),
            2 => left.wrapping_mul(right),
            3 | 5 => {
                let shift = 64 - WIDTH * 8;
                let left = ((left << shift) as i64 >> shift) as i128;
                let right = ((right << shift) as i64 >> shift) as i128;
                let limit = 1i128 << (WIDTH * 8 - 1);
                let result = if OP == 3 { left + right } else { left - right };
                result.clamp(-limit, limit - 1) as u64
            }
            4 => left.saturating_add(right).min(u64::MAX >> (64 - WIDTH * 8)),
            6 => left.saturating_sub(right),
            _ => unreachable!("unsupported integer SIMD kernel"),
        };
        out[start..start + WIDTH].copy_from_slice(&result.to_le_bytes()[..WIDTH]);
    }
    out
}

macro_rules! binary {
    ($name:ident, $width:literal, $op:literal, $arm:expr, $x86:ident) => {
        #[inline(always)]
        pub(super) fn $name(a: &[u8; 16], b: &[u8; 16]) -> [u8; 16] {
            #[cfg(all(target_endian = "little", target_arch = "aarch64"))]
            {
                // Both local arrays contain 16 initialized bytes; NEON's
                // byte-vector accesses do not require guest alignment.
                unsafe {
                    let left = vld1q_u8(a.as_ptr());
                    let right = vld1q_u8(b.as_ptr());
                    let result = ($arm)(left, right);
                    let mut out = [0u8; 16];
                    vst1q_u8(out.as_mut_ptr(), result);
                    out
                }
            }
            #[cfg(all(target_endian = "little", target_arch = "x86_64"))]
            {
                // SSE2 is baseline on x86-64. No optional target features or
                // aligned loads are assumed.
                unsafe {
                    let left = _mm_loadu_si128(a.as_ptr().cast());
                    let right = _mm_loadu_si128(b.as_ptr().cast());
                    let result = $x86(left, right);
                    let mut out = [0u8; 16];
                    _mm_storeu_si128(out.as_mut_ptr().cast(), result);
                    out
                }
            }
            #[cfg(all(target_arch = "wasm32", target_feature = "simd128"))]
            {
                use core::arch::wasm32 as wasm;
                unsafe {
                    // Each input and output holds exactly sixteen bytes;
                    // standard WASM vector loads/stores permit unaligned access.
                    let left = wasm::v128_load(a.as_ptr().cast());
                    let right = wasm::v128_load(b.as_ptr().cast());
                    // Width and operation are macro literals, so optimized
                    // builds select a single standard SIMD instruction.
                    let result = match ($width, $op) {
                        (1, 0) => wasm::i8x16_add(left, right),
                        (1, 1) => wasm::i8x16_sub(left, right),
                        (1, 3) => wasm::i8x16_add_sat(left, right),
                        (1, 4) => wasm::u8x16_add_sat(left, right),
                        (1, 5) => wasm::i8x16_sub_sat(left, right),
                        (1, 6) => wasm::u8x16_sub_sat(left, right),
                        (2, 0) => wasm::i16x8_add(left, right),
                        (2, 1) => wasm::i16x8_sub(left, right),
                        (2, 2) => wasm::i16x8_mul(left, right),
                        (2, 3) => wasm::i16x8_add_sat(left, right),
                        (2, 4) => wasm::u16x8_add_sat(left, right),
                        (2, 5) => wasm::i16x8_sub_sat(left, right),
                        (2, 6) => wasm::u16x8_sub_sat(left, right),
                        (4, 0) => wasm::i32x4_add(left, right),
                        (4, 1) => wasm::i32x4_sub(left, right),
                        (8, 0) => wasm::i64x2_add(left, right),
                        (8, 1) => wasm::i64x2_sub(left, right),
                        _ => return scalar::<$width, $op>(a, b),
                    };
                    let mut out = [0u8; 16];
                    wasm::v128_store(out.as_mut_ptr().cast(), result);
                    out
                }
            }
            #[cfg(not(any(
                all(target_endian = "little", any(target_arch = "aarch64", target_arch = "x86_64")),
                all(target_arch = "wasm32", target_feature = "simd128")
            )))]
            {
                scalar::<$width, $op>(a, b)
            }
        }
    };
}


/// Standard i8x16.eq: every lane is exactly 0xff or 0x00.
#[inline(always)]
pub(super) fn i8x16_eq(a: &[u8; 16], b: &[u8; 16]) -> [u8; 16] {
    #[cfg(all(target_endian = "little", target_arch = "aarch64"))]
    unsafe {
        // Local arrays hold sixteen initialized bytes; no guest alignment is assumed.
        let mut out = [0u8; 16];
        vst1q_u8(out.as_mut_ptr(), vceqq_u8(vld1q_u8(a.as_ptr()), vld1q_u8(b.as_ptr())));
        out
    }
    #[cfg(all(target_endian = "little", target_arch = "x86_64"))]
    unsafe {
        // SSE2 is baseline on x86-64; both local vector loads are unaligned.
        let mut out = [0u8; 16];
        _mm_storeu_si128(out.as_mut_ptr().cast(), _mm_cmpeq_epi8(_mm_loadu_si128(a.as_ptr().cast()), _mm_loadu_si128(b.as_ptr().cast())));
        out
    }
    #[cfg(not(all(target_endian = "little", any(target_arch = "aarch64", target_arch = "x86_64"))))]
    {
        let mut out = [0u8; 16];
        for lane in 0..16 { out[lane] = if a[lane] == b[lane] { 0xff } else { 0 }; }
        out
    }
}

/// Standard i8x16.ne: every lane is exactly 0xff or 0x00.
#[inline(always)]
pub(super) fn i8x16_ne(a: &[u8; 16], b: &[u8; 16]) -> [u8; 16] {
    #[cfg(all(target_endian = "little", target_arch = "aarch64"))]
    unsafe {
        // All accesses stay within initialized sixteen-byte local arrays.
        let mut out = [0u8; 16];
        vst1q_u8(out.as_mut_ptr(), vmvnq_u8(vceqq_u8(vld1q_u8(a.as_ptr()), vld1q_u8(b.as_ptr()))));
        out
    }
    #[cfg(all(target_endian = "little", target_arch = "x86_64"))]
    unsafe {
        // SSE2 is baseline on x86-64, and these loads permit unaligned arrays.
        let mut out = [0u8; 16];
        let equal = _mm_cmpeq_epi8(_mm_loadu_si128(a.as_ptr().cast()), _mm_loadu_si128(b.as_ptr().cast()));
        _mm_storeu_si128(out.as_mut_ptr().cast(), _mm_xor_si128(equal, _mm_set1_epi8(-1)));
        out
    }
    #[cfg(all(target_arch = "wasm32", target_feature = "simd128"))]
    unsafe {
        use core::arch::wasm32 as wasm;
        let mut out = [0u8; 16];
        wasm::v128_store(out.as_mut_ptr().cast(), wasm::i8x16_ne(wasm::v128_load(a.as_ptr().cast()), wasm::v128_load(b.as_ptr().cast())));
        out
    }
    #[cfg(not(any(
        all(target_endian = "little", any(target_arch = "aarch64", target_arch = "x86_64")),
        all(target_arch = "wasm32", target_feature = "simd128")
    )))]
    {
        let mut out = [0u8; 16];
        for lane in 0..16 { out[lane] = if a[lane] != b[lane] { 0xff } else { 0 }; }
        out
    }
}

binary!(i8x16_add, 1, 0, |a, b| vaddq_u8(a, b), _mm_add_epi8);
binary!(i8x16_sub, 1, 1, |a, b| vsubq_u8(a, b), _mm_sub_epi8);
binary!(i8x16_add_sat_s, 1, 3, |a, b| vreinterpretq_u8_s8(vqaddq_s8(vreinterpretq_s8_u8(a), vreinterpretq_s8_u8(b))), _mm_adds_epi8);
binary!(i8x16_add_sat_u, 1, 4, |a, b| vqaddq_u8(a, b), _mm_adds_epu8);
binary!(i8x16_sub_sat_s, 1, 5, |a, b| vreinterpretq_u8_s8(vqsubq_s8(vreinterpretq_s8_u8(a), vreinterpretq_s8_u8(b))), _mm_subs_epi8);
binary!(i8x16_sub_sat_u, 1, 6, |a, b| vqsubq_u8(a, b), _mm_subs_epu8);
binary!(i16x8_add, 2, 0, |a, b| vreinterpretq_u8_u16(vaddq_u16(vreinterpretq_u16_u8(a), vreinterpretq_u16_u8(b))), _mm_add_epi16);
binary!(i16x8_sub, 2, 1, |a, b| vreinterpretq_u8_u16(vsubq_u16(vreinterpretq_u16_u8(a), vreinterpretq_u16_u8(b))), _mm_sub_epi16);
binary!(i16x8_add_sat_s, 2, 3, |a, b| vreinterpretq_u8_s16(vqaddq_s16(vreinterpretq_s16_u8(a), vreinterpretq_s16_u8(b))), _mm_adds_epi16);
binary!(i16x8_add_sat_u, 2, 4, |a, b| vreinterpretq_u8_u16(vqaddq_u16(vreinterpretq_u16_u8(a), vreinterpretq_u16_u8(b))), _mm_adds_epu16);
binary!(i16x8_sub_sat_s, 2, 5, |a, b| vreinterpretq_u8_s16(vqsubq_s16(vreinterpretq_s16_u8(a), vreinterpretq_s16_u8(b))), _mm_subs_epi16);
binary!(i16x8_sub_sat_u, 2, 6, |a, b| vreinterpretq_u8_u16(vqsubq_u16(vreinterpretq_u16_u8(a), vreinterpretq_u16_u8(b))), _mm_subs_epu16);
binary!(i16x8_mul, 2, 2, |a, b| vreinterpretq_u8_u16(vmulq_u16(vreinterpretq_u16_u8(a), vreinterpretq_u16_u8(b))), _mm_mullo_epi16);
binary!(i32x4_add, 4, 0, |a, b| vreinterpretq_u8_u32(vaddq_u32(vreinterpretq_u32_u8(a), vreinterpretq_u32_u8(b))), _mm_add_epi32);
binary!(i32x4_sub, 4, 1, |a, b| vreinterpretq_u8_u32(vsubq_u32(vreinterpretq_u32_u8(a), vreinterpretq_u32_u8(b))), _mm_sub_epi32);
binary!(i64x2_add, 8, 0, |a, b| vreinterpretq_u8_u64(vaddq_u64(vreinterpretq_u64_u8(a), vreinterpretq_u64_u8(b))), _mm_add_epi64);
binary!(i64x2_sub, 8, 1, |a, b| vreinterpretq_u8_u64(vsubq_u64(vreinterpretq_u64_u8(a), vreinterpretq_u64_u8(b))), _mm_sub_epi64);

/// Extract the sign bits of eight independent 16-bit lanes in lane order.
#[inline(always)]
pub(super) fn i16x8_bitmask(bytes: &[u8; 16]) -> i32 {
    let compact = |word: u64| {
        // Extract the four lane sign bits, then place them in bits 48..51.
        // Each base-65536 column contains distinct powers of two, so the
        // multiplication cannot carry into the selected four output bits.
        let signs = (word >> 15) & 0x0001_0001_0001_0001;
        ((signs.wrapping_mul(0x0001_0002_0004_0008) >> 48) & 15) as i32
    };
    let low = u64::from_le_bytes(bytes[..8].try_into().unwrap());
    let high = u64::from_le_bytes(bytes[8..].try_into().unwrap());
    compact(low) | (compact(high) << 4)
}

/// Rounded unsigned average, computed in eight independent 16-bit lanes.
#[inline(always)]
pub(super) fn i16x8_avgr_u(a: &[u8; 16], b: &[u8; 16]) -> [u8; 16] {
    i16x8_avgr_u_scalar(a, b)
}

// Retained only to reproduce the rejected hardware-kernel experiment.
#[cfg(test)]
#[inline(always)]
fn i16x8_avgr_u_intrinsic(a: &[u8; 16], b: &[u8; 16]) -> [u8; 16] {
    #[cfg(all(target_endian = "little", target_arch = "aarch64"))]
    unsafe {
        // Baseline NEON. Both inputs and output hold sixteen initialized
        // bytes; byte loads need no stronger alignment than these arrays.
        let left = vreinterpretq_u16_u8(vld1q_u8(a.as_ptr()));
        let right = vreinterpretq_u16_u8(vld1q_u8(b.as_ptr()));
        let mut out = [0u8; 16];
        vst1q_u8(out.as_mut_ptr(), vreinterpretq_u8_u16(vrhaddq_u16(left, right)));
        out
    }
    #[cfg(all(target_endian = "little", target_arch = "x86_64"))]
    unsafe {
        // SSE2 is baseline on x86-64, including Linux and Windows. PAVGW
        // computes ceil((left + right) / 2) without overflowing the lanes.
        let left = _mm_loadu_si128(a.as_ptr().cast());
        let right = _mm_loadu_si128(b.as_ptr().cast());
        let mut out = [0u8; 16];
        _mm_storeu_si128(out.as_mut_ptr().cast(), _mm_avg_epu16(left, right));
        out
    }
    #[cfg(not(all(target_endian = "little", any(target_arch = "aarch64", target_arch = "x86_64"))))]
    { i16x8_avgr_u_scalar(a, b) }
}

#[inline(always)]
fn i16x8_avgr_u_scalar(a: &[u8; 16], b: &[u8; 16]) -> [u8; 16] {
    let a = u128::from_le_bytes(*a);
    let b = u128::from_le_bytes(*b);
    // Clear each lane's low bit before shifting, preventing bits crossing
    // lane boundaries. Per lane, (a | b) >= ((a ^ b) >> 1), so subtraction
    // cannot borrow from an adjacent lane. This is exactly ceil((a+b)/2)
    // without the overflowing lane addition or any memory access/unsafe.
    const SHIFT_MASK: u128 = (u128::MAX / 65535) * 65534;
    ((a | b) - (((a ^ b) & SHIFT_MASK) >> 1)).to_le_bytes()
}

#[cfg(test)]
mod rounded_average_tests {
    use super::i16x8_avgr_u;

    fn check(a: [u8; 16], b: [u8; 16]) {
        let actual = i16x8_avgr_u(&a, &b);
        assert_eq!(actual, super::i16x8_avgr_u_intrinsic(&a, &b));
        for lane in 0..8 {
            let offset = lane * 2;
            let left = u16::from_le_bytes([a[offset], a[offset + 1]]) as u32;
            let right = u16::from_le_bytes([b[offset], b[offset + 1]]) as u32;
            let result = u16::from_le_bytes([actual[offset], actual[offset + 1]]);
            assert_eq!(result, ((left + right + 1) / 2) as u16);
        }
    }

    #[test]
    fn all_u16_values_against_boundary_lanes() {
        for value in 0..=u16::MAX {
            let a = std::array::from_fn(|i| value.to_le_bytes()[i % 2]);
            let edges = [0u16, 1, 2, 0x7fff, 0x8000, 0xfffe, 0xffff, !value];
            let b = std::array::from_fn(|i| edges[i / 2].to_le_bytes()[i % 2]);
            check(a, b);
            check(b, a);
        }
    }

    #[test]
    fn mixed_lanes_do_not_carry_or_borrow() {
        let mut state = 0x12345678u32;
        for _ in 0..65536 {
            let mut next = || {
                state = state.wrapping_mul(1664525).wrapping_add(1013904223);
                (state >> 24) as u8
            };
            check(std::array::from_fn(|_| next()), std::array::from_fn(|_| next()));
        }
    }

    #[test]
    #[ignore = "paired exact SIMD rounded-average performance measurement"]
    fn rounded_average_kernel_benchmark() {
        use std::{hint::black_box, time::Instant};
        type Kernel = fn(&[u8; 16], &[u8; 16]) -> [u8; 16];
        fn time(kernel: Kernel, a: &[u8; 16], b: &[u8; 16]) -> u128 {
            let kernel = black_box(kernel);
            let start = Instant::now();
            for _ in 0..1_000_000 {
                black_box(kernel(black_box(a), black_box(b)));
            }
            start.elapsed().as_nanos()
        }
        let a = [0, 0, 255, 255, 0, 128, 1, 0, 255, 127, 254, 255, 91, 171, 3, 193];
        let b = [255, 255, 255, 255, 255, 127, 0, 0, 0, 128, 1, 0, 239, 2, 44, 205];
        assert_eq!(super::i16x8_avgr_u_intrinsic(&a, &b), super::i16x8_avgr_u_scalar(&a, &b));
        let mut ratios = [0.0f64; 11];
        for (pair, ratio) in ratios.iter_mut().enumerate() {
            let (old, new) = if pair % 2 == 0 {
                (time(super::i16x8_avgr_u_scalar, &a, &b), time(super::i16x8_avgr_u_intrinsic, &a, &b))
            } else {
                let new = time(super::i16x8_avgr_u_intrinsic, &a, &b);
                (time(super::i16x8_avgr_u_scalar, &a, &b), new)
            };
            *ratio = old as f64 / new as f64;
        }
        ratios.sort_by(f64::total_cmp);
        eprintln!("i16x8.avgr_u scalar/native median {:.3}x; pairs {ratios:?}", ratios[5]);
    }
}

#[cfg(test)]
mod wide_bitmask_tests {
    use super::i16x8_bitmask;

    #[test]
    fn every_sign_mask_preserves_lane_order() {
        for mask in 0..256 {
            for payload in [0u8, 1, 0x55, 0x7f, 0xff] {
                let bytes = std::array::from_fn(|i| {
                    if i % 2 == 0 { payload } else {
                        (payload & 127) | ((((mask >> (i / 2)) & 1) as u8) << 7)
                    }
                });
                assert_eq!(i16x8_bitmask(&bytes), mask);
            }
        }
    }

    #[test]
    fn arbitrary_bits_match_signed_lane_oracle() {
        let mut state = 7u32;
        for _ in 0..65536 {
            let bytes = std::array::from_fn(|_| {
                state = state.wrapping_mul(1664525).wrapping_add(1013904223);
                (state >> 24) as u8
            });
            let mut expected = 0;
            for lane in 0..8 {
                if i16::from_le_bytes([bytes[2 * lane], bytes[2 * lane + 1]]) < 0 {
                    expected |= 1 << lane;
                }
            }
            assert_eq!(i16x8_bitmask(&bytes), expected);
        }
    }
}

#[inline(always)]
fn swizzle_scalar(source: &[u8; 16], indices: &[u8; 16]) -> [u8; 16] {
    #[cfg(all(target_arch = "x86_64", target_endian = "little"))]
    if std::is_x86_feature_detected!("ssse3") {
        // Detection is cached by std. Never execute SSSE3 on an unsupported CPU.
        return unsafe { swizzle_ssse3(source, indices) };
    }
    let mut out = [0u8; 16];
    for lane in 0..16 {
        let index = indices[lane] as usize;
        out[lane] = if index < 16 { source[index] } else { 0 };
    }
    out
}

#[inline(always)]
pub(super) fn i8x16_swizzle(source: &[u8; 16], indices: &[u8; 16]) -> [u8; 16] {
    #[cfg(all(target_endian = "little", target_arch = "aarch64"))]
    {
        // Baseline NEON table lookup zeroes every index outside 0..16,
        // including indices with their high bit set. Only initialized local
        // vector bytes are accessed; no guest pointer or alignment is assumed.
        unsafe {
            let table = vld1q_u8(source.as_ptr());
            let selectors = vld1q_u8(indices.as_ptr());
            let mut out = [0u8; 16];
            vst1q_u8(out.as_mut_ptr(), vqtbl1q_u8(table, selectors));
            out
        }
    }
    #[cfg(not(all(target_endian = "little", target_arch = "aarch64")))]
    {
        // In particular, do not assume optional SSSE3 on x86-64.
        swizzle_scalar(source, indices)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    type SatKernel = fn(&[u8; 16], &[u8; 16]) -> [u8; 16];

    #[test]
    fn swizzle_matches_every_selector_and_mixed_vectors() {
        let source = std::array::from_fn(|lane| lane as u8 * 13 + 7);
        for index in 0..=255u16 {
            let indices = [index as u8; 16];
            let expected = if index < 16 { [source[index as usize]; 16] } else { [0; 16] };
            assert_eq!(i8x16_swizzle(&source, &indices), expected, "index {index}");
            assert_eq!(swizzle_scalar(&source, &indices), expected);
        }
        let mut state = 0x6a09_e667_f3bc_c909u64;
        for _ in 0..4096 {
            let mut source = [0u8; 16];
            let mut indices = [0u8; 16];
            for byte in source.iter_mut().chain(indices.iter_mut()) {
                state = state.wrapping_mul(6364136223846793005).wrapping_add(1);
                *byte = (state >> 32) as u8;
            }
            assert_eq!(i8x16_swizzle(&source, &indices), swizzle_scalar(&source, &indices));
        }
    }

    #[test]
    #[ignore = "paired release swizzle-kernel performance measurement"]
    fn swizzle_kernel_benchmark() {
        use std::hint::black_box;
        use std::time::Instant;
        fn time(kernel: SatKernel, source: &[u8; 16], indices: &[u8; 16]) -> f64 {
            let start = Instant::now();
            for _ in 0..2_000_000 {
                black_box(black_box(kernel)(black_box(source), black_box(indices)));
            }
            start.elapsed().as_secs_f64()
        }
        let source = std::array::from_fn(|lane| lane as u8 * 13 + 7);
        for (name, indices) in [
            ("in-range", [15, 14, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1, 0]),
            ("mixed", [0, 15, 16, 127, 128, 255, 1, 14, 17, 3, 0, 240, 13, 12, 11, 10]),
            ("out-of-range", [255; 16]),
        ] {
            assert_eq!(i8x16_swizzle(&source, &indices), swizzle_scalar(&source, &indices));
            let mut ratios = [0.0f64; 5];
            for (pair, ratio) in ratios.iter_mut().enumerate() {
                let (old, new) = if pair % 2 == 0 {
                    (time(swizzle_scalar, &source, &indices), time(i8x16_swizzle, &source, &indices))
                } else {
                    let new = time(i8x16_swizzle, &source, &indices);
                    (time(swizzle_scalar, &source, &indices), new)
                };
                *ratio = old / new;
            }
            ratios.sort_by(f64::total_cmp);
            eprintln!("swizzle {name}: scalar/native median {:.3}x; pairs {ratios:?}", ratios[2]);
        }
    }

    fn saturation_reference<const WIDTH: usize, const OP: u8>(
        a: &[u8; 16], b: &[u8; 16],
    ) -> [u8; 16] {
        let mut out = [0u8; 16];
        for start in (0..16).step_by(WIDTH) {
            if WIDTH == 1 {
                out[start] = match OP {
                    3 => (a[start] as i8).saturating_add(b[start] as i8) as u8,
                    4 => a[start].saturating_add(b[start]),
                    5 => (a[start] as i8).saturating_sub(b[start] as i8) as u8,
                    6 => a[start].saturating_sub(b[start]),
                    _ => unreachable!(),
                };
            } else {
                let left = u16::from_le_bytes([a[start], a[start + 1]]);
                let right = u16::from_le_bytes([b[start], b[start + 1]]);
                let result = match OP {
                    3 => (left as i16).saturating_add(right as i16) as u16,
                    4 => left.saturating_add(right),
                    5 => (left as i16).saturating_sub(right as i16) as u16,
                    6 => left.saturating_sub(right),
                    _ => unreachable!(),
                };
                out[start..start + 2].copy_from_slice(&result.to_le_bytes());
            }
        }
        out
    }

    fn saturation_kernels() -> [(&'static str, SatKernel, SatKernel, SatKernel); 8] {
        [
            ("i8.add_s", i8x16_add_sat_s, scalar::<1, 3>, saturation_reference::<1, 3>),
            ("i8.add_u", i8x16_add_sat_u, scalar::<1, 4>, saturation_reference::<1, 4>),
            ("i8.sub_s", i8x16_sub_sat_s, scalar::<1, 5>, saturation_reference::<1, 5>),
            ("i8.sub_u", i8x16_sub_sat_u, scalar::<1, 6>, saturation_reference::<1, 6>),
            ("i16.add_s", i16x8_add_sat_s, scalar::<2, 3>, saturation_reference::<2, 3>),
            ("i16.add_u", i16x8_add_sat_u, scalar::<2, 4>, saturation_reference::<2, 4>),
            ("i16.sub_s", i16x8_sub_sat_s, scalar::<2, 5>, saturation_reference::<2, 5>),
            ("i16.sub_u", i16x8_sub_sat_u, scalar::<2, 6>, saturation_reference::<2, 6>),
        ]
    }

    #[test]
    fn saturation_matches_scalar_and_lane_reference() {
        let kernels = saturation_kernels();
        // Every possible byte-lane operand pair, including signed boundaries.
        for left in 0..=255u16 {
            for right in 0..=255u16 {
                let a = [left as u8; 16];
                let b = [right as u8; 16];
                for &(name, native, portable, reference) in &kernels[..4] {
                    let expected = reference(&a, &b);
                    assert_eq!(native(&a, &b), expected, "{name}: {left}, {right}");
                    assert_eq!(portable(&a, &b), expected, "portable {name}");
                }
            }
        }
        let edges = [0u16, 1, 2, 0x7ffe, 0x7fff, 0x8000, 0x8001, 0xfffe, 0xffff];
        for left in edges {
            for right in edges {
                let mut a = [0u8; 16];
                let mut b = [0u8; 16];
                for lane in 0..8 {
                    a[lane * 2..lane * 2 + 2].copy_from_slice(&left.to_le_bytes());
                    b[lane * 2..lane * 2 + 2].copy_from_slice(&right.to_le_bytes());
                }
                for &(name, native, portable, reference) in &kernels[4..] {
                    let expected = reference(&a, &b);
                    assert_eq!(native(&a, &b), expected, "{name}: {left}, {right}");
                    assert_eq!(portable(&a, &b), expected, "portable {name}");
                }
            }
        }
        let mut state = 0x6a09_e667_f3bc_c909u64;
        for _ in 0..4096 {
            let mut a = [0u8; 16];
            let mut b = [0u8; 16];
            for byte in a.iter_mut().chain(b.iter_mut()) {
                state = state.wrapping_mul(6364136223846793005).wrapping_add(1);
                *byte = (state >> 32) as u8;
            }
            for &(name, native, portable, reference) in &kernels {
                let expected = reference(&a, &b);
                assert_eq!(native(&a, &b), expected, "{name}");
                assert_eq!(portable(&a, &b), expected, "portable {name}");
            }
        }
    }

    #[test]
    #[ignore = "paired release performance measurement"]
    fn saturation_kernel_benchmark() {
        use std::hint::black_box;
        use std::time::Instant;
        fn time(kernel: SatKernel, a: &[u8; 16], b: &[u8; 16]) -> f64 {
            let start = Instant::now();
            for _ in 0..200_000 {
                black_box(black_box(kernel)(black_box(a), black_box(b)));
            }
            start.elapsed().as_secs_f64()
        }
        let a = [0x7f, 0x80, 0xff, 0, 1, 0x7e, 0xfe, 2, 0x80, 1, 0xff, 0xff, 0, 0, 0x7f, 0xff];
        let b = [1, 0xff, 2, 0xff, 0x80, 4, 8, 0x7f, 0x80, 0xff, 1, 0, 0xff, 0xff, 0x80, 1];
        for (name, native, _, reference) in saturation_kernels() {
            assert_eq!(native(&a, &b), reference(&a, &b), "{name}");
            let mut ratios = [0.0f64; 5];
            for (pair, ratio) in ratios.iter_mut().enumerate() {
                let (old, new) = if pair % 2 == 0 {
                    (time(reference, &a, &b), time(native, &a, &b))
                } else {
                    let new = time(native, &a, &b);
                    (time(reference, &a, &b), new)
                };
                *ratio = old / new;
            }
            ratios.sort_by(f64::total_cmp);
            eprintln!("{name}: reference/native median {:.3}x; pairs {ratios:?}", ratios[2]);
        }
    }

    #[test]
    fn native_integer_kernels_match_wrapping_lanes() {
        type Kernel = fn(&[u8; 16], &[u8; 16]) -> [u8; 16];
        let kernels: &[(Kernel, Kernel)] = &[
            (i8x16_add, scalar::<1, 0>),
            (i8x16_sub, scalar::<1, 1>),
            (i16x8_add, scalar::<2, 0>),
            (i16x8_sub, scalar::<2, 1>),
            (i16x8_mul, scalar::<2, 2>),
            (i32x4_add, scalar::<4, 0>),
            (i32x4_sub, scalar::<4, 1>),
            (i64x2_add, scalar::<8, 0>),
            (i64x2_sub, scalar::<8, 1>),
        ];
        let mut state = 0x6a09_e667_f3bc_c909u64;
        for sample in 0..2048 {
            let mut a = [0u8; 16];
            let mut b = [0u8; 16];
            for byte in a.iter_mut().chain(b.iter_mut()) {
                state = state.wrapping_mul(6364136223846793005).wrapping_add(1);
                *byte = (state >> 32) as u8;
            }
            if sample < 16 {
                a.fill(sample as u8 * 17);
                b.fill(255 - sample as u8 * 17);
            }
            for &(native, reference) in kernels {
                assert_eq!(native(&a, &b), reference(&a, &b), "sample {sample}");
            }
        }
    }
}

#[cfg(any(test, not(any(
    all(target_endian = "little", target_arch = "aarch64"),
    all(target_endian = "little", target_arch = "x86_64")
))))]
#[inline(always)]
fn bitmask_i8_scalar(bytes: &[u8; 16]) -> i32 {
    let mut mask = 0;
    for (lane, byte) in bytes.iter().enumerate() {
        mask |= ((byte >> 7) as i32) << lane;
    }
    mask
}

#[inline(always)]
pub(super) fn i8x16_bitmask(bytes: &[u8; 16]) -> i32 {
    #[cfg(all(target_endian = "little", target_arch = "aarch64"))]
    {
        use std::arch::aarch64::*;
        unsafe {
            let weights = [1u8, 2, 4, 8, 16, 32, 64, 128, 1, 2, 4, 8, 16, 32, 64, 128];
            let signs = vshrq_n_s8::<7>(vreinterpretq_s8_u8(vld1q_u8(bytes.as_ptr())));
            let bits = vandq_u8(vreinterpretq_u8_s8(signs), vld1q_u8(weights.as_ptr()));
            // Each half sums distinct powers of two, so its sum fits u8.
            i32::from(vaddv_u8(vget_low_u8(bits)))
                | (i32::from(vaddv_u8(vget_high_u8(bits))) << 8)
        }
    }
    #[cfg(all(target_endian = "little", target_arch = "x86_64"))]
    {
        use std::arch::x86_64::*;
        unsafe { _mm_movemask_epi8(_mm_loadu_si128(bytes.as_ptr().cast())) }
    }
    #[cfg(all(target_arch = "wasm32", target_feature = "simd128"))]
    {
        // The build explicitly requires standard SIMD128 support. v128.load
        // permits unaligned addresses and reads exactly these sixteen bytes.
        unsafe {
            i32::from(core::arch::wasm32::i8x16_bitmask(core::arch::wasm32::v128_load(
                bytes.as_ptr().cast(),
            )))
        }
    }
    #[cfg(not(any(
        all(target_endian = "little", target_arch = "aarch64"),
        all(target_endian = "little", target_arch = "x86_64"),
        all(target_arch = "wasm32", target_feature = "simd128")
    )))]
    { bitmask_i8_scalar(bytes) }
}

#[cfg(test)]
mod bitmask_tests {
    use super::*;

    #[test]
    fn all_i8_sign_masks_match_lane_order() {
        for mask in 0..=u16::MAX {
            let mut bytes = [0u8; 16];
            for (lane, byte) in bytes.iter_mut().enumerate() {
                *byte = (((mask >> lane) & 1) as u8) << 7
                    | ((lane * 11 + mask as usize) & 127) as u8;
            }
            assert_eq!(bitmask_i8_scalar(&bytes), i32::from(mask));
            assert_eq!(i8x16_bitmask(&bytes), i32::from(mask));
        }
    }

    #[test]
    #[ignore = "alternating loaded-machine bitmask kernel measurement"]
    fn bitmask_kernel_loaded_benchmark() {
        let input = [0, 255, 127, 128, 91, 171, 3, 193, 239, 2, 44, 205, 78, 148, 113, 129];
        let run = |kernel: fn(&[u8; 16]) -> i32| {
            let start = std::time::Instant::now();
            for _ in 0..5_000_000 {
                std::hint::black_box(kernel(std::hint::black_box(&input)));
            }
            start.elapsed().as_nanos()
        };
        run(bitmask_i8_scalar);
        run(i8x16_bitmask);
        let mut before = Vec::with_capacity(21);
        let mut after = Vec::with_capacity(21);
        let mut ratios = Vec::with_capacity(21);
        for round in 0..21 {
            let (a, b) = if round % 2 == 0 { (run(bitmask_i8_scalar), run(i8x16_bitmask)) }
                else { let b = run(i8x16_bitmask); (run(bitmask_i8_scalar), b) };
            before.push(a); after.push(b); ratios.push(a as f64 / b as f64);
        }
        before.sort_unstable(); after.sort_unstable(); ratios.sort_by(f64::total_cmp);
        eprintln!("i8x16.bitmask scalar_p50_ns={} native_p50_ns={} scalar_p95_ns={} native_p95_ns={} paired_ratio_p10={:.3} paired_ratio_p50={:.3} paired_ratio_p90={:.3}", before[10], after[10], before[19], after[19], ratios[2], ratios[10], ratios[18]);
    }
}

#[cfg(all(target_arch = "x86_64", target_endian = "little"))]
#[target_feature(enable = "ssse3")]
unsafe fn swizzle_ssse3(source: &[u8; 16], indices: &[u8; 16]) -> [u8; 16] {
    use core::arch::x86_64::{_mm_adds_epu8, _mm_loadu_si128, _mm_set1_epi8,
        _mm_shuffle_epi8, _mm_storeu_si128};
    let source = unsafe { _mm_loadu_si128(source.as_ptr().cast()) };
    let indices = unsafe { _mm_loadu_si128(indices.as_ptr().cast()) };
    // PSHUFB alone would wrap selectors 16..127. Saturating addition of 0x70
    // keeps the low nibble for 0..15, and sets the zeroing high bit for 16..255.
    let controls = _mm_adds_epu8(indices, _mm_set1_epi8(0x70));
    let result = _mm_shuffle_epi8(source, controls);
    let mut bytes = [0u8; 16];
    unsafe { _mm_storeu_si128(bytes.as_mut_ptr().cast(), result) };
    bytes
}

#[cfg(test)]
mod swizzle_backend_boundary_tests {
    use super::*;

    #[test]
    fn biased_shuffle_controls_match_wasm_for_every_selector() {
        let source = std::array::from_fn(|i| (i * 13 + 7) as u8);
        for index in 0u16..=255 {
            let indices = [index as u8; 16];
            let expected = [if index < 16 { source[index as usize] } else { 0 }; 16];
            let controls = indices.map(|i| i.saturating_add(0x70));
            let shuffle_model = controls.map(|i| if i & 0x80 != 0 { 0 } else { source[(i & 15) as usize] });
            assert_eq!(shuffle_model, expected);
            // On SSSE3 x86 this also exercises the new hardware path. Elsewhere
            // it verifies the retained scalar fallback against the same oracle.
            assert_eq!(swizzle_scalar(&source, &indices), expected);
        }
    }

    #[test]
    fn swizzle_backend_mixed_lanes_preserve_zeroing_and_selection() {
        for round in 0usize..4096 {
            let source = std::array::from_fn(|i| (round.wrapping_mul(29) + i * 71) as u8);
            let indices = std::array::from_fn(|i| (round.wrapping_mul(17) + i * 43) as u8);
            let expected = std::array::from_fn(|i| source.get(indices[i] as usize).copied().unwrap_or(0));
            assert_eq!(swizzle_scalar(&source, &indices), expected);
        }
    }
}

#[cfg(all(test, target_arch = "x86_64", target_endian = "little"))]
mod ssse3_swizzle_execution_tests {
    use super::*;

    #[test]
    fn ssse3_intrinsic_executes_all_selector_boundaries() {
        if !std::is_x86_feature_detected!("ssse3") {
            eprintln!("SSSE3 unavailable: intrinsic execution skipped");
            return;
        }
        let source = std::array::from_fn(|i| (i * 13 + 7) as u8);
        for index in 0u16..=255 {
            let indices = [index as u8; 16];
            let expected = [if index < 16 { source[index as usize] } else { 0 }; 16];
            assert_eq!(unsafe { swizzle_ssse3(&source, &indices) }, expected);
        }
        eprintln!("Executed SSSE3 intrinsic kernel for all 256 selectors");
    }

    #[test]
    #[ignore = "paired x86 kernel measurement; label Rosetta runs separately"]
    fn ssse3_swizzle_paired_benchmark() {
        use std::hint::black_box;
        use std::time::Instant;
        if !std::is_x86_feature_detected!("ssse3") {
            eprintln!("SSSE3 unavailable: native measurement skipped");
            return;
        }
        fn previous(source: &[u8; 16], indices: &[u8; 16]) -> [u8; 16] {
            std::array::from_fn(|i| source.get(indices[i] as usize).copied().unwrap_or(0))
        }
        type Kernel = fn(&[u8; 16], &[u8; 16]) -> [u8; 16];
        let old = black_box(previous as Kernel);
        let new = black_box(swizzle_scalar as Kernel);
        let source = std::array::from_fn(|i| (i * 37 + 9) as u8);
        for (label, indices) in [
            ("valid", std::array::from_fn(|i| (15 - i) as u8)),
            ("mixed", std::array::from_fn(|i| (i * 17) as u8)),
            ("invalid", [16u8; 16]),
        ] {
            assert_eq!(old(&source, &indices), new(&source, &indices));
            let mut before = Vec::new();
            let mut after = Vec::new();
            for pair in 0..23 {
                for candidate in if pair % 2 == 0 { [false, true] } else { [true, false] } {
                    let kernel = if candidate { new } else { old };
                    let start = Instant::now();
                    for _ in 0..100_000 { black_box(kernel(black_box(&source), black_box(&indices))); }
                    let duration = start.elapsed().as_nanos();
                    if pair >= 2 { if candidate { after.push(duration); } else { before.push(duration); } }
                }
            }
            before.sort_unstable(); after.sort_unstable();
            eprintln!("x86-swizzle {label} old_p50_ns={} new_p50_ns={} old_p95_ns={} new_p95_ns={} ratio={:.3}",
                before[10], after[10], before[19], after[19], before[10] as f64 / after[10] as f64);
        }
    }
}

#[cfg(test)]
mod equality_tests {
fn scalar(a: &[u8; 16], b: &[u8; 16]) -> [u8; 16] {
    let mut out = [0; 16];
    for lane in 0..16 {
        out[lane] = if a[lane] == b[lane] { 0xff } else { 0 };
    }
    out
}

#[test]
fn every_equality_mask_retains_exact_lane_bytes() {
    for mask in 0..=u16::MAX {
        let a = std::array::from_fn(|lane| (mask as u8).wrapping_add(lane as u8 * 13));
        let b = std::array::from_fn(|lane| {
            if mask & (1 << lane) != 0 { a[lane] } else { a[lane].wrapping_add(1) }
        });
        let expected = scalar(&a, &b);
        assert_eq!(super::i8x16_eq(&a, &b), expected);
        assert_eq!(super::i8x16_ne(&a, &b), expected.map(|lane| !lane));
        assert_eq!(super::i8x16_bitmask(&expected), i32::from(mask));
    }
}

#[test]
fn all_byte_pairs_match_at_every_lane() {
    for left in 0..=u8::MAX {
        for right in 0..=u8::MAX {
            let a = std::array::from_fn(|lane| left.wrapping_add(lane as u8 * 13));
            let b = std::array::from_fn(|lane| right.wrapping_add(lane as u8 * 13));
            assert_eq!(super::i8x16_eq(&a, &b), scalar(&a, &b));
            assert_eq!(super::i8x16_ne(&a, &b), scalar(&a, &b).map(|lane| !lane));
        }
    }
}

}

#[inline(always)]
pub(super) fn i16x8_eq(a: &[u8; 16], b: &[u8; 16]) -> [u8; 16] {
    #[cfg(all(target_endian = "little", target_arch = "aarch64"))]
    unsafe {
        // Sixteen initialized bytes; unaligned baseline NEON loads.
        let mut out = [0u8; 16];
        vst1q_u8(out.as_mut_ptr(), vreinterpretq_u8_u16(vceqq_s16(vld1q_s16(a.as_ptr().cast()), vld1q_s16(b.as_ptr().cast()))));
        out
    }
    #[cfg(all(target_endian = "little", target_arch = "x86_64"))]
    unsafe {
        // SSE2 is baseline on x86-64, including unaligned loads/stores.
        let mut out = [0u8; 16];
        _mm_storeu_si128(out.as_mut_ptr().cast(), _mm_cmpeq_epi16(_mm_loadu_si128(a.as_ptr().cast()), _mm_loadu_si128(b.as_ptr().cast())));
        out
    }
    #[cfg(not(all(target_endian = "little", any(target_arch = "aarch64", target_arch = "x86_64"))))]
    {
        let mut out = [0u8; 16];
        for lane in 0..8 {
            let start = lane * 2;
            if a[start..start + 2] == b[start..start + 2] { out[start..start + 2].fill(0xff); }
        }
        out
    }
}

#[inline(always)]
pub(super) fn i32x4_eq(a: &[u8; 16], b: &[u8; 16]) -> [u8; 16] {
    #[cfg(all(target_endian = "little", target_arch = "aarch64"))]
    unsafe {
        // Sixteen initialized bytes; unaligned baseline NEON loads.
        let mut out = [0u8; 16];
        vst1q_u8(out.as_mut_ptr(), vreinterpretq_u8_u32(vceqq_s32(vld1q_s32(a.as_ptr().cast()), vld1q_s32(b.as_ptr().cast()))));
        out
    }
    #[cfg(all(target_endian = "little", target_arch = "x86_64"))]
    unsafe {
        // SSE2 is baseline on x86-64, including unaligned loads/stores.
        let mut out = [0u8; 16];
        _mm_storeu_si128(out.as_mut_ptr().cast(), _mm_cmpeq_epi32(_mm_loadu_si128(a.as_ptr().cast()), _mm_loadu_si128(b.as_ptr().cast())));
        out
    }
    #[cfg(not(all(target_endian = "little", any(target_arch = "aarch64", target_arch = "x86_64"))))]
    {
        let mut out = [0u8; 16];
        for lane in 0..4 {
            let start = lane * 4;
            if a[start..start + 4] == b[start..start + 4] { out[start..start + 4].fill(0xff); }
        }
        out
    }
}
#[cfg(test)]
mod wide_equality_tests {
use super::*;
fn oracle(a:&[u8;16], b:&[u8;16], width:usize)->[u8;16] {
 let mut out=[0u8;16];
 for (lane,(left,right)) in a.chunks_exact(width).zip(b.chunks_exact(width)).enumerate() {
  if left==right {out[lane*width..(lane+1)*width].fill(255);}
 }
 out
}
#[test]
fn exact_masks_and_full_lane_bits() {
 for width in [2,4] {
  let compare=if width==2 {i16x8_eq} else {i32x4_eq};
  for mask in 0..(1usize<<(16/width)) {
   let a=std::array::from_fn(|i|(i*17+128) as u8);let mut b=a;
   for lane in 0..16/width {if mask&(1<<lane)==0 {b[lane*width+width-1]^=128;}}
   assert_eq!(compare(&a,&b),oracle(&a,&b,width));
  }
  let mut random=0x12345678u32;
  for iteration in 0..65536 {
   let mut a=[0u8;16];let mut b=[0u8;16];
   for i in 0..16 {random=random.wrapping_mul(1664525).wrapping_add(1013904223);a[i]=(random>>24) as u8;random=random.wrapping_mul(1664525).wrapping_add(1013904223);b[i]=(random>>24) as u8;}
   let lane=iteration%(16/width);b[lane*width..(lane+1)*width].copy_from_slice(&a[lane*width..(lane+1)*width]);
   assert_eq!(compare(&a,&b),oracle(&a,&b,width));
   assert_eq!(compare(&a,&a),[255;16]);
  }
 }
}

}

/// Return exactly 0 or 1 according to whether all eight integer lanes are nonzero.
#[cfg(test)]
#[inline(always)]
pub(super) fn i16x8_all_true(bytes: &[u8; 16]) -> i32 {
    #[cfg(all(target_arch = "aarch64", target_feature = "neon"))]
    {
        if u16::from_le_bytes([bytes[0], bytes[1]]) == 0 {
            return 0;
        }
        use core::arch::aarch64::{vld1q_u8, vminvq_u16, vreinterpretq_u16_u8};
        // SAFETY: NEON is enabled for this target. The borrowed array is live
        // and provides all 16 readable bytes. Loading u8 lanes requires no
        // stronger alignment than the input; reinterpretation is register-only.
        // Endianness cannot change whether either byte pair represents zero.
        unsafe {
            (vminvq_u16(vreinterpretq_u16_u8(vld1q_u8(bytes.as_ptr()))) != 0) as i32
        }
    }
    #[cfg(not(all(target_arch = "aarch64", target_feature = "neon")))]
    {
        (0..8).all(|lane| {
            u16::from_le_bytes([bytes[lane * 2], bytes[lane * 2 + 1]]) != 0
        }) as i32
    }
}

#[cfg(test)]
mod all_true_reduction_tests {
    use super::i16x8_all_true;

    fn oracle(bytes: &[u8; 16]) -> i32 {
        bytes.chunks_exact(2).all(|lane| lane[0] != 0 || lane[1] != 0) as i32
    }

    #[test]
    fn i16_all_true_exhausts_each_lane_value_and_zero_mask() {
        for lane in 0..8 {
            for value in 0..=u16::MAX {
                let mut bytes = [1u8; 16];
                bytes[lane * 2..lane * 2 + 2].copy_from_slice(&value.to_le_bytes());
                assert_eq!(i16x8_all_true(&bytes), oracle(&bytes));
            }
        }
        for mask in 0..256 {
            for value in [1u16, 128, 256, 32768, 65535] {
                let mut bytes = [0u8; 16];
                for lane in 0..8 {
                    if mask & (1 << lane) == 0 {
                        bytes[lane * 2..lane * 2 + 2]
                            .copy_from_slice(&value.to_le_bytes());
                    }
                }
                assert_eq!(i16x8_all_true(&bytes), oracle(&bytes));
            }
        }
    }

    #[test]
    fn i16_all_true_mixed_random_vectors_match_independent_oracle() {
        let mut seed = 0x41b8a531u32;
        for round in 0..65536 {
            let mut bytes = [0u8; 16];
            for byte in &mut bytes {
                seed ^= seed << 13;
                seed ^= seed >> 17;
                seed ^= seed << 5;
                *byte = seed as u8;
            }
            if round % 3 != 0 {
                let offset = (round % 8) * 2;
                bytes[offset..offset + 2].fill(0);
            }
            assert_eq!(i16x8_all_true(&bytes), oracle(&bytes));
        }
    }
}
