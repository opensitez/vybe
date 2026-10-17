include!("../src/emitter/sdl_runtime.rs");

fn original_rgba_scale(
    input: &[u8], output: &mut [u8], width: usize, rows: usize, pitch: usize,
    output_width: usize, output_rows: usize, output_pitch: usize,
) -> Option<()> {
    if width == 0 || rows == 0 || output_width == 0 || output_rows == 0 { return None; }
    let source_bytes = width.checked_mul(4)?;
    let output_bytes = output_width.checked_mul(4)?;
    if pitch < source_bytes || output_pitch < output_bytes { return None; }
    let source_len = (rows - 1).checked_mul(pitch)?.checked_add(source_bytes)?;
    let output_len = (output_rows - 1).checked_mul(output_pitch)?.checked_add(output_bytes)?;
    if input.len() < source_len || output.len() < output_len { return None; }
    if output_rows > rows && width != output_width {
        return scale_rgba_enlarged_rows(input, output, width, rows, pitch,
            output_width, output_rows, output_pitch);
    }
    let (xs, xr) = (width / output_width, width % output_width);
    let (ys, yr) = (rows / output_rows, rows % output_rows);
    let (xt, yt) = (output_width - xr, output_rows - yr);
    let (mut source_y, mut fraction_y) = (0, 0);
    for y in 0..output_rows {
        let row = &input[source_y * pitch..source_y * pitch + source_bytes];
        let destination = &mut output[y * output_pitch..y * output_pitch + output_bytes];
        if width == output_width { destination.copy_from_slice(row); }
        else {
            let (mut source_x, mut fraction_x) = (0, 0);
            for pixel in destination.chunks_exact_mut(4) {
                pixel.copy_from_slice(&row[source_x * 4..source_x * 4 + 4]);
                source_x += xs;
                if fraction_x >= xt { fraction_x -= xt; source_x += 1; }
                else { fraction_x += xr; }
            }
        }
        source_y += ys;
        if fraction_y >= yt { fraction_y -= yt; source_y += 1; }
        else { fraction_y += yr; }
    }
    Some(())
}

fn double_truecolor_frame(
    input: &[u8], output: &mut [u8], width: usize, rows: usize, pitch: usize,
    output_rows: usize, output_pitch: usize,
) -> Option<()> {
    if width == 0 || rows == 0 || output_rows == 0 { return None; }
    let source_bytes = width.checked_mul(4)?;
    let output_bytes = width.checked_mul(8)?;
    if pitch < source_bytes || output_pitch < output_bytes { return None; }
    let source_span = (rows - 1).checked_mul(pitch)?.checked_add(source_bytes)?;
    let output_span = (output_rows - 1).checked_mul(output_pitch)?.checked_add(output_bytes)?;
    if input.len() < source_span || output.len() < output_span { return None; }
    let (ys, yr) = (rows / output_rows, rows % output_rows);
    let yt = output_rows - yr;
    let (mut source_y, mut fraction_y) = (0, 0);
    let mut previous_y = None;
    for y in 0..output_rows {
        if previous_y == Some(source_y) {
            let previous = (y - 1) * output_pitch;
            output.copy_within(previous..previous + output_bytes, y * output_pitch);
        } else {
            let source = &input[source_y * pitch..source_y * pitch + source_bytes];
            let destination = &mut output[y * output_pitch..y * output_pitch + output_bytes];
            for (pixel, pair) in source.chunks_exact(4).zip(destination.chunks_exact_mut(8)) {
                pair[..4].copy_from_slice(pixel);
                pair[4..].copy_from_slice(pixel);
            }
        }
        previous_y = Some(source_y);
        source_y += ys;
        if fraction_y >= yt { fraction_y -= yt; source_y += 1; }
        else { fraction_y += yr; }
    }
    Some(())
}

#[test]
fn truecolor_double_preserves_every_byte_padding_and_vertical_mapping() {
    for width in 1..18usize {
        for rows in 1..10usize {
            for output_rows in 1..20usize {
                for padding in [0, 13] {
                    let pitch = width * 4 + padding;
                    let output_pitch = width * 8 + padding;
                    let input: Vec<u8> = (0..(rows - 1) * pitch + width * 4)
                        .map(|index| (index * 73 + 91) as u8).collect();
                    let mut expected = vec![0xa5; (output_rows - 1) * output_pitch + width * 8 + 19];
                    for y in 0..output_rows {
                        for x in 0..width * 2 {
                            let source = (y * rows / output_rows) * pitch + (x / 2) * 4;
                            let destination = y * output_pitch + x * 4;
                            expected[destination..destination + 4].copy_from_slice(&input[source..source + 4]);
                        }
                    }
                    let mut actual = vec![0xa5; expected.len()];
                    assert_eq!(double_truecolor_frame(&input, &mut actual, width, rows, pitch,
                        output_rows, output_pitch), Some(()));
                    assert_eq!(actual, expected);
                    actual.fill(0xa5);
                    assert_eq!(scale_rgba_bytes(&input, &mut actual, width, rows, pitch,
                        width * 2, output_rows, output_pitch), Some(()));
                    assert_eq!(actual, expected);
                }
            }
        }
    }
}

#[test]
#[ignore = "paired production exact-2x truecolor scaler versus original implementation; not Doom FPS"]
fn paired_truecolor_double_frame_benchmark() {
    use std::hint::black_box;
    use std::time::Instant;
    for (width, rows, output_rows) in [(320usize, 200usize, 100usize), (320, 200, 200),
        (320, 200, 400), (640, 480, 960)] {
        let pitch = width * 4 + 13;
        let output_pitch = width * 8 + 13;
        let input: Vec<u8> = (0..(rows - 1) * pitch + width * 4)
            .map(|index| (index * 73 + 91) as u8).collect();
        let mut samples = [Vec::new(), Vec::new()];
        for round in 0..19 {
            let order = if round % 2 == 0 { [0, 1] } else { [1, 0] };
            for path in order {
                let mut output = vec![0xa5; (output_rows - 1) * output_pitch + width * 8];
                let start = Instant::now();
                for _ in 0..20 {
                    if path == 0 {
                        original_rgba_scale(black_box(&input), black_box(&mut output), width, rows, pitch,
                            width * 2, output_rows, output_pitch).unwrap();
                    } else {
                        scale_rgba_bytes(black_box(&input), black_box(&mut output), width, rows, pitch,
                            width * 2, output_rows, output_pitch).unwrap();
                    }
                }
                let elapsed = start.elapsed().as_secs_f64();
                let mut expected = vec![0xa5; output.len()];
                original_rgba_scale(&input, &mut expected, width, rows, pitch,
                    width * 2, output_rows, output_pitch).unwrap();
                assert_eq!(output, expected);
                if round >= 2 { samples[path].push(elapsed); }
            }
        }
        for sample in &mut samples { sample.sort_by(f64::total_cmp); }
        eprintln!("truecolor-double-frame {width}x{rows} -> {}x{output_rows} production_ms={:.3} double_ms={:.3} production_over_double={:.3}",
            width * 2, samples[0][8] * 1000.0, samples[1][8] * 1000.0, samples[0][8] / samples[1][8]);
    }
}
