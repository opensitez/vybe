include!("sdl_truecolor_double_frame.rs");

fn integer_scale<const FACTOR: usize>(
    input: &[u8], output: &mut [u8], width: usize, rows: usize,
    pitch: usize, output_rows: usize, output_pitch: usize,
) -> Option<()> {
    if !(3..=4).contains(&FACTOR) || width == 0 || rows == 0 || output_rows == 0 {
        return None;
    }
    let source_bytes = width.checked_mul(4)?;
    let destination_bytes = source_bytes.checked_mul(FACTOR)?;
    if pitch < source_bytes || output_pitch < destination_bytes { return None; }
    let source_end = (rows - 1).checked_mul(pitch)?.checked_add(source_bytes)?;
    let destination_end = (output_rows - 1).checked_mul(output_pitch)?.checked_add(destination_bytes)?;
    if source_end > input.len() || destination_end > output.len() { return None; }
    let mut previous = None;
    for y in 0..output_rows {
        let source_y = ((y as u128 * rows as u128) / output_rows as u128) as usize;
        let destination_start = y * output_pitch;
        if previous == Some(source_y) {
            let previous_start = destination_start - output_pitch;
            output.copy_within(previous_start..previous_start + destination_bytes, destination_start);
        } else {
            let source_start = source_y * pitch;
            let source = &input[source_start..source_start + source_bytes];
            let destination = &mut output[destination_start..destination_start + destination_bytes];
            for (pixel, expanded) in source.chunks_exact(4).zip(destination.chunks_exact_mut(4 * FACTOR)) {
                for copy in expanded.chunks_exact_mut(4) { copy.copy_from_slice(pixel); }
            }
        }
        previous = Some(source_y);
    }
    Some(())
}

#[test]
fn integer_scaling_matches_current_scaler_and_preserves_guards() {
    for width in 1..=7 {
        for rows in 1..=5 {
            for output_rows in 1..=11 {
                let pitch = width * 4 + 3;
                let source: Vec<u8> = (0..rows * pitch).map(|i| (i * 73 + 19) as u8).collect();
                for factor in [3, 4] {
                    let destination_pitch = width * factor * 4 + 5;
                    let mut actual = vec![0x91; output_rows * destination_pitch + 17];
                    let mut expected = actual.clone();
                    original_rgba_scale(&source, &mut expected, width, rows, pitch,
                        width * factor, output_rows, destination_pitch).unwrap();
                    let mut production = actual.clone();
                    scale_rgba_bytes(&source, &mut production, width, rows, pitch,
                        width * factor, output_rows, destination_pitch).unwrap();
                    assert_eq!(production, expected);
                    let result = if factor == 3 {
                        integer_scale::<3>(&source, &mut actual, width, rows, pitch, output_rows, destination_pitch)
                    } else {
                        integer_scale::<4>(&source, &mut actual, width, rows, pitch, output_rows, destination_pitch)
                    };
                    assert_eq!(result, Some(()));
                    assert_eq!(actual, expected, "{width}x{rows}, factor={factor}, output_rows={output_rows}");
                }
            }
        }
    }
}

#[test]
fn integer_scaling_rejects_invalid_ranges_without_writes() {
    let source = [1u8; 64];
    let mut output = [0x91; 128];
    for (width, rows, pitch, output_rows, output_pitch) in [
        (0, 1, 4, 1, 12), (1, 0, 4, 1, 12), (1, 1, 4, 0, 12),
        (1, 1, 3, 1, 12), (1, 1, 4, 1, 11), (2, 9, 8, 1, 24),
        (2, 1, 8, 6, 24), (usize::MAX, 1, 4, 1, 12),
        (1, usize::MAX, 4, 1, 12), (1, 1, 4, usize::MAX, 12),
    ] {
        assert_eq!(integer_scale::<3>(&source, &mut output, width, rows, pitch, output_rows, output_pitch), None);
        assert_eq!(output, [0x91; 128]);
    }
}

#[test]
#[ignore = "paired integer-scaling prototype benchmark; not Doom FPS"]
fn paired_integer_scaling_benchmark() {
    use std::{hint::black_box, time::Instant};
    let width = 320;
    let rows = 200;
    let pitch = width * 4 + 13;
    let source: Vec<u8> = (0..rows * pitch).map(|i| (i * 37) as u8).collect();
    for factor in [3, 4] {
        let output_width = width * factor;
        let output_rows = rows * factor;
        let output_pitch = output_width * 4 + 7;
        let mut output = vec![0x91; output_rows * output_pitch];
        let mut expected = output.clone();
        original_rgba_scale(&source, &mut expected, width, rows, pitch, output_width, output_rows, output_pitch).unwrap();
        let mut times = [Vec::new(), Vec::new()];
        for round in 0..19 {
            for position in 0..2 {
                let path = (round + position) % 2;
                let start = Instant::now();
                for _ in 0..20 {
                    if path == 0 {
                        original_rgba_scale(black_box(&source), black_box(&mut output), width, rows, pitch, output_width, output_rows, output_pitch).unwrap();
                    } else if factor == 3 {
                        scale_rgba_bytes(black_box(&source), black_box(&mut output), width, rows, pitch, output_width, output_rows, output_pitch).unwrap();
                    } else {
                        scale_rgba_bytes(black_box(&source), black_box(&mut output), width, rows, pitch, output_width, output_rows, output_pitch).unwrap();
                    }
                    black_box(&output);
                }
                let elapsed = start.elapsed().as_secs_f64() * 1000.0;
                assert_eq!(output, expected);
                if round >= 2 { times[path].push(elapsed); }
            }
        }
        for samples in &mut times { samples.sort_by(f64::total_cmp); }
        eprintln!("integer-{factor}x current_ms={:.3} prototype_ms={:.3} speedup={:.3}", times[0][8], times[1][8], times[0][8] / times[1][8]);
    }
}
