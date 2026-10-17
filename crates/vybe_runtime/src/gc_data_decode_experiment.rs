//! Test-only candidate: exact-length collection versus push-based GC decoding.
use crate::Value;
use std::{hint::black_box, time::Instant};

fn pushed(bytes: &[u8], wide: bool) -> Vec<Value> {
    let mut result = Vec::with_capacity(bytes.len() / if wide { 8 } else { 4 });
    if wide {
        for lane in bytes.chunks_exact(8) {
            result.push(Value::I64(i64::from_le_bytes(lane.try_into().unwrap())));
        }
    } else {
        for lane in bytes.chunks_exact(4) {
            result.push(Value::I32(i32::from_le_bytes(lane.try_into().unwrap())));
        }
    }
    result
}

fn collected(bytes: &[u8], wide: bool) -> Vec<Value> {
    if wide {
        bytes.chunks_exact(8)
            .map(|lane| Value::I64(i64::from_le_bytes(lane.try_into().unwrap())))
            .collect()
    } else {
        bytes.chunks_exact(4)
            .map(|lane| Value::I32(i32::from_le_bytes(lane.try_into().unwrap())))
            .collect()
    }
}

#[test]
fn exact_length_decode_matches_push_for_signed_values_and_unaligned_slices() {
    let bytes: Vec<u8> = (0usize..8193).map(|i| (i.wrapping_mul(37) + 129) as u8).collect();
    for wide in [false, true] {
        let width = if wide { 8 } else { 4 };
        for offset in [0, 1, 3] {
            for count in [0, 1, 2, 17, 1023] {
                let bytes = &bytes[offset..offset + count * width];
                assert_eq!(collected(bytes, wide), pushed(bytes, wide));
            }
        }
    }
}

#[test]
#[ignore = "GC numeric data decoding allocation experiment, not production"]
fn exact_length_gc_data_decode_benchmark() {
    for wide in [false, true] {
        for count in [16usize, 1024, 32768] {
            let bytes: Vec<u8> = (0..count * if wide { 8 } else { 4 })
                .map(|i| (i.wrapping_mul(37) + 129) as u8).collect();
            let iterations = (262144 / count).max(32);
            let mut samples = [Vec::new(), Vec::new()];
            for pair in 0..9 {
                for step in 0..2 {
                    let index = (pair + step) % 2;
                    let clock = Instant::now();
                    for _ in 0..iterations {
                        let values = if index == 0 { pushed(black_box(&bytes), wide) }
                            else { collected(black_box(&bytes), wide) };
                        black_box(values);
                    }
                    if pair >= 2 { samples[index].push(clock.elapsed().as_secs_f64()); }
                }
            }
            for sample in &mut samples { sample.sort_by(f64::total_cmp); }
            eprintln!("gc-data wide={wide} count={count} iterations={iterations} pushed_ms={:.3} collected_ms={:.3} pushed_over_collected={:.3}",
                samples[0][3] * 1000.0, samples[1][3] * 1000.0, samples[0][3] / samples[1][3]);
        }
    }
}
