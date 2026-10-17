//! Standalone WASM execution gate for the production byte sign-mask backend.
//! Build with rustc --crate-type=cdylib --target wasm32-unknown-unknown
//! -C opt-level=3 -C target-feature=+simd128, then invoke _start in a WASM engine.
#![allow(dead_code)]

#[path = "../../src/simd_kernels.rs"]
mod kernels;

#[no_mangle]
pub extern "C" fn _start() {
    for mask in 0..=u16::MAX {
        let mut storage = [0u8; 17];
        for lane in 0..16 {
            storage[lane + 1] = (((mask >> lane) & 1) as u8) << 7
                | ((usize::from(mask) * 11 + lane * 13) & 127) as u8;
        }
        let bytes: &[u8; 16] = (&storage[1..]).try_into().unwrap();
        let actual = kernels::i8x16_bitmask(std::hint::black_box(bytes));
        if actual != i32::from(mask) {
            core::arch::wasm32::unreachable();
        }
    }
}
