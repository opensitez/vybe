//! Standalone execution gate for the actual vector-copy byte kernel.

#[path = "../../src/v128_copy.rs"]
mod kernel;

fn check(condition: bool) {
    if !condition {
        core::arch::wasm32::unreachable();
    }
}

#[no_mangle]
pub extern "C" fn _start() {
    let mut initial = [0u8; 256];
    for (index, byte) in initial.iter_mut().enumerate() {
        *byte = (index * 37 + 11) as u8;
    }
    for src in 0..=33 {
        for dst in 0..=33 {
            for bytes in [16usize, 32, 64, 128] {
                let mut expected = initial;
                let mut last = [0u8; 16];
                for offset in (0..bytes).step_by(16) {
                    for lane in 0..16 {
                        last[lane] = expected[src + offset + lane];
                    }
                    for lane in 0..16 {
                        expected[dst + offset + lane] = last[lane];
                    }
                }
                let mut actual = initial;
                check(kernel::copy_forward_vectors(&mut actual, src, dst, bytes) == Some(last));
                check(actual == expected);
            }
        }
    }
    for (src, dst, bytes) in [
        (0usize, 0usize, 0usize), (0, 0, 1), (0, 0, 17),
        (usize::MAX, 0, 16), (0, usize::MAX, 16),
        (0, 0, usize::MAX - 15), (241, 0, 16), (0, 241, 16),
    ] {
        let mut actual = initial;
        check(kernel::copy_forward_vectors(&mut actual, src, dst, bytes).is_none());
        check(actual == initial);
    }
}
