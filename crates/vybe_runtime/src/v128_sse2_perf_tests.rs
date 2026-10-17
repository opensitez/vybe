use super::copy_forward_vectors;
use std::{hint::black_box, time::Instant};

fn scalar(memory: &mut [u8], src: usize, dst: usize, bytes: usize) -> Option<[u8; 16]> {
    if bytes == 0 || bytes % 16 != 0 { return None; }
    if src.checked_add(bytes)? > memory.len() || dst.checked_add(bytes)? > memory.len() { return None; }
    let mut last = [0; 16];
    for offset in (0..bytes).step_by(16) {
        last.copy_from_slice(&memory[src + offset..src + offset + 16]);
        memory[dst + offset..dst + offset + 16].copy_from_slice(&last);
    }
    Some(last)
}

#[test]
#[ignore = "paired SSE2 carry versus forward scalar snapshot loop"]
fn sse2_carry_benchmark() {
    for gap in [9usize, 12, 15] {
        for bytes in [128usize, 512, 4096, 65536] {
            let initial: Vec<u8> = (0..bytes + gap + 37)
                .map(|i| (i as u8).wrapping_mul(37).wrapping_add(11)).collect();
            let methods = [scalar as fn(&mut [u8], usize, usize, usize) -> Option<[u8; 16]>, copy_forward_vectors];
            let mut samples = [Vec::new(), Vec::new()];
            let rounds = (262144 / bytes).max(2);
            for trial in 0..21 {
                let mut reference = None;
                for order in 0..2 {
                    let which = (trial + order) % 2;
                    let mut memory = initial.clone();
                    let start = Instant::now();
                    let mut last = None;
                    for _ in 0..rounds {
                        last = methods[which](black_box(&mut memory), 3, 3 + gap, black_box(bytes));
                        black_box(last);
                    }
                    let ns = start.elapsed().as_nanos() as f64 / rounds as f64;
                    if let Some((expected, expected_last)) = &reference {
                        assert_eq!(&memory, expected);
                        assert_eq!(last, *expected_last);
                    } else { reference = Some((memory, last)); }
                    samples[which].push(ns);
                }
            }
            for sample in &mut samples { sample.sort_by(f64::total_cmp); }
            println!("sse2-carry gap={gap} bytes={bytes} scalar_ns={:.2} current_ns={:.2} speedup={:.3}",
                samples[0][10], samples[1][10], samples[0][10] / samples[1][10]);
        }
    }
}
