//! Standalone WASM gate for the actual production integer SIMD kernels.
#![allow(dead_code)]

#[path = "../../src/simd_kernels.rs"]
mod kernels;

fn reference(a: &[u8; 16], b: &[u8; 16], width: usize, op: u8) -> [u8; 16] {
    let mut out = [0u8; 16];
    for start in (0..16).step_by(width) {
        let mut x = [0u8; 8];
        let mut y = [0u8; 8];
        x[..width].copy_from_slice(&a[start..start + width]);
        y[..width].copy_from_slice(&b[start..start + width]);
        let left = u64::from_le_bytes(x);
        let right = u64::from_le_bytes(y);
        let result = match (width, op) {
            (_, 0) => left.wrapping_add(right),
            (_, 1) => left.wrapping_sub(right),
            (_, 2) => left.wrapping_mul(right),
            (1, 3) => (left as i8).saturating_add(right as i8) as u8 as u64,
            (1, 4) => (left as u8).saturating_add(right as u8) as u64,
            (1, 5) => (left as i8).saturating_sub(right as i8) as u8 as u64,
            (1, 6) => (left as u8).saturating_sub(right as u8) as u64,
            (2, 3) => (left as i16).saturating_add(right as i16) as u16 as u64,
            (2, 4) => (left as u16).saturating_add(right as u16) as u64,
            (2, 5) => (left as i16).saturating_sub(right as i16) as u16 as u64,
            (2, 6) => (left as u16).saturating_sub(right as u16) as u64,
            _ => core::arch::wasm32::unreachable(),
        };
        out[start..start + width].copy_from_slice(&result.to_le_bytes()[..width]);
    }
    out
}

fn check(a: &[u8; 16], b: &[u8; 16]) {
    macro_rules! check_kernel {
        ($name:ident, $width:literal, $op:literal) => {
            if kernels::$name(std::hint::black_box(a), std::hint::black_box(b))
                != reference(a, b, $width, $op)
            {
                core::arch::wasm32::unreachable();
            }
        };
    }
    check_kernel!(i8x16_add, 1, 0);
    check_kernel!(i8x16_sub, 1, 1);
    check_kernel!(i8x16_add_sat_s, 1, 3);
    check_kernel!(i8x16_add_sat_u, 1, 4);
    check_kernel!(i8x16_sub_sat_s, 1, 5);
    check_kernel!(i8x16_sub_sat_u, 1, 6);
    check_kernel!(i16x8_add, 2, 0);
    check_kernel!(i16x8_sub, 2, 1);
    check_kernel!(i16x8_mul, 2, 2);
    check_kernel!(i16x8_add_sat_s, 2, 3);
    check_kernel!(i16x8_add_sat_u, 2, 4);
    check_kernel!(i16x8_sub_sat_s, 2, 5);
    check_kernel!(i16x8_sub_sat_u, 2, 6);
    check_kernel!(i32x4_add, 4, 0);
    check_kernel!(i32x4_sub, 4, 1);
    check_kernel!(i64x2_add, 8, 0);
    check_kernel!(i64x2_sub, 8, 1);
}

#[no_mangle]
pub extern "C" fn _start() {
    for left in 0..=u8::MAX {
        for right in 0..=u8::MAX {
            check(&[left; 16], &[right; 16]);
        }
    }
    let mut state = 0x6a09_e667_f3bc_c909u64;
    for _ in 0..4096 {
        let mut left = [0u8; 17];
        let mut right = [0u8; 17];
        for byte in left.iter_mut().chain(right.iter_mut()) {
            state = state.wrapping_mul(6364136223846793005).wrapping_add(1);
            *byte = (state >> 32) as u8;
        }
        check((&left[1..]).try_into().unwrap(), (&right[1..]).try_into().unwrap());
    }
}
