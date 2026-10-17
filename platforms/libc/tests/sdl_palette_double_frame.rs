include!("sdl_palette_double_row.rs");

fn double_palette_frame(
    input: &[u8], output: &mut [u8], width: usize, rows: usize, pitch: usize,
    output_rows: usize, output_pitch: usize, colors: &[[u8; 4]],
) -> Option<()> {
    if width == 0 || rows == 0 || output_rows == 0 || colors.len() < 256 { return None; }
    let output_bytes = width.checked_mul(8)?;
    if pitch < width || output_pitch < output_bytes { return None; }
    let source_span = (rows - 1).checked_mul(pitch)?.checked_add(width)?;
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
            let source = &input[source_y * pitch..source_y * pitch + width];
            let destination = &mut output[y * output_pitch..y * output_pitch + output_bytes];
            double_palette_row(source, destination, colors)?;
        }
        previous_y = Some(source_y);
        source_y += ys;
        if fraction_y >= yt { fraction_y -= yt; source_y += 1; }
        else { fraction_y += yr; }
    }
    Some(())
}

#[test]
fn double_palette_frames_preserve_nearest_mapping_padding_and_guards() {
    let colors: Vec<[u8; 4]> = (0..256usize)
        .map(|index| [index as u8, (index * 17) as u8, (index * 37) as u8, (index * 53) as u8]).collect();
    for width in 1..18usize {
        for rows in 1..10usize {
            for output_rows in 1..20usize {
                for padding in [0, 13] {
                    let pitch = width + padding;
                    let output_pitch = width * 8 + padding;
                    let input: Vec<u8> = (0..(rows - 1) * pitch + width)
                        .map(|index| (index * 73 + 91) as u8).collect();
                    let mut expected = vec![0xa5; (output_rows - 1) * output_pitch + width * 8 + 19];
                    let mut actual = expected.clone();
                    assert_eq!(palette_scale_reference(&input, &mut expected, width, rows, pitch,
                        width * 2, output_rows, output_pitch, &colors), Some(()));
                    assert_eq!(double_palette_frame(&input, &mut actual, width, rows, pitch,
                        output_rows, output_pitch, &colors), Some(()));
                    assert_eq!(actual, expected);
                    actual.fill(0xa5);
                    assert_eq!(scale_palette_bytes(&input, &mut actual, width, rows, pitch,
                        width * 2, output_rows, output_pitch, &colors), Some(()));
                    assert_eq!(actual, expected);
                }
            }
        }
    }
    let mut output = [0xa5; 31];
    assert_eq!(double_palette_frame(&[0; 4], &mut output, 2, 2, 2, 2, 16, &colors), None);
    assert_eq!(output, [0xa5; 31]);
}

#[test]
#[ignore = "production exact-2x palette full-frame comparison versus original scaler; not Doom FPS"]
fn paired_palette_double_frame_benchmark() {
    use std::hint::black_box;
    use std::time::Instant;
    let colors: Vec<[u8; 4]> = (0..256usize)
        .map(|index| [index as u8, (index * 17) as u8, (index * 37) as u8, (index * 53) as u8]).collect();
    for (width, rows, output_rows) in [(320usize, 200usize, 100usize), (320, 200, 200),
        (320, 200, 400), (640, 480, 960)] {
        let pitch = width + 13;
        let output_pitch = width * 8 + 13;
        let input: Vec<u8> = (0..(rows - 1) * pitch + width)
            .map(|index| (index * 73 + 91) as u8).collect();
        let mut samples = [Vec::new(), Vec::new()];
        for round in 0..19 {
            let order = if round % 2 == 0 { [0, 1] } else { [1, 0] };
            for path in order {
                let mut output = vec![0xa5; (output_rows - 1) * output_pitch + width * 8];
                let start = Instant::now();
                for _ in 0..20 {
                    if path == 0 {
                        palette_scale_reference(black_box(&input), black_box(&mut output), width, rows, pitch,
                            width * 2, output_rows, output_pitch, black_box(&colors)).unwrap();
                    } else {
                        scale_palette_bytes(black_box(&input), black_box(&mut output), width, rows, pitch,
                            width * 2, output_rows, output_pitch, black_box(&colors)).unwrap();
                    }
                }
                let elapsed = start.elapsed().as_secs_f64();
                let mut expected = vec![0xa5; output.len()];
                palette_scale_reference(&input, &mut expected, width, rows, pitch,
                    width * 2, output_rows, output_pitch, &colors).unwrap();
                assert_eq!(output, expected);
                if round >= 2 { samples[path].push(elapsed); }
            }
        }
        for sample in &mut samples { sample.sort_by(f64::total_cmp); }
        eprintln!("palette-double-frame {width}x{rows} -> {}x{output_rows} generic_ms={:.3} double_ms={:.3} generic_over_double={:.3}",
            width * 2, samples[0][8] * 1000.0, samples[1][8] * 1000.0, samples[0][8] / samples[1][8]);
    }
}
