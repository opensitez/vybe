include!("../src/emitter/sdl_runtime.rs");

fn double_palette_row(input: &[u8], output: &mut [u8], colors: &[[u8; 4]]) -> Option<()> {
    frame_kernels::double_palette_row(input, output, colors)
}

fn palette_scale_reference(
    input: &[u8], output: &mut [u8], width: usize, rows: usize, pitch: usize,
    output_width: usize, output_rows: usize, output_pitch: usize, colors: &[[u8; 4]],
) -> Option<()> {
    if width == 0 || rows == 0 || output_width == 0 || output_rows == 0 || colors.len() < 256 { return None; }
    let output_bytes = output_width.checked_mul(4)?;
    if pitch < width || output_pitch < output_bytes { return None; }
    let source_span = (rows - 1).checked_mul(pitch)?.checked_add(width)?;
    let output_span = (output_rows - 1).checked_mul(output_pitch)?.checked_add(output_bytes)?;
    if input.len() < source_span || output.len() < output_span { return None; }
    let (xs, xr) = (width / output_width, width % output_width);
    let (ys, yr) = (rows / output_rows, rows % output_rows);
    let (xt, yt) = (output_width - xr, output_rows - yr);
    let (mut source_y, mut fraction_y) = (0, 0);
    let mut previous_y = None;
    for y in 0..output_rows {
        if previous_y == Some(source_y) {
            let previous = (y - 1) * output_pitch;
            output.copy_within(previous..previous + output_bytes, y * output_pitch);
        } else {
            let row = &input[source_y * pitch..source_y * pitch + width];
            let destination = &mut output[y * output_pitch..y * output_pitch + output_bytes];
            let (mut source_x, mut fraction_x) = (0, 0);
            for pixel in destination.chunks_exact_mut(4) {
                pixel.copy_from_slice(&colors[usize::from(row[source_x])]);
                source_x += xs;
                if fraction_x >= xt { fraction_x -= xt; source_x += 1; }
                else { fraction_x += xr; }
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
fn palette_double_row_matches_checked_scaler_and_preserves_tail() {
    let colors: Vec<[u8; 4]> = (0..256usize)
        .map(|index| [index as u8, (index * 17) as u8, (index * 37) as u8, (index * 53) as u8]).collect();
    for width in 1..=513usize {
        let input: Vec<u8> = (0..width).map(|index| (index * 73 + 91) as u8).collect();
        let mut expected = vec![0xa5; width * 8 + 19];
        let mut actual = expected.clone();
        assert_eq!(palette_scale_reference(&input, &mut expected, width, 1, width,
            width * 2, 1, width * 8, &colors), Some(()));
        assert_eq!(double_palette_row(&input, &mut actual, &colors), Some(()));
        assert_eq!(actual, expected);
    }
    let mut output = [0xa5; 15];
    assert_eq!(double_palette_row(&[0, 1], &mut output, &colors), None);
    assert_eq!(output, [0xa5; 15]);
    assert_eq!(double_palette_row(&[0], &mut output, &colors[..255]), None);
    assert_eq!(output, [0xa5; 15]);
}

#[test]
#[ignore = "exact-2x palette row experiment; not full-frame or Doom speedup"]
fn paired_palette_double_row_benchmark() {
    use std::hint::black_box;
    use std::time::Instant;
    let colors: Vec<[u8; 4]> = (0..256usize)
        .map(|index| [index as u8, (index * 17) as u8, (index * 37) as u8, (index * 53) as u8]).collect();
    for width in [320usize, 640, 1280] {
        let input: Vec<u8> = (0..width).map(|index| (index * 73 + 91) as u8).collect();
        let mut samples = [Vec::new(), Vec::new()];
        for round in 0..19 {
            let order = if round % 2 == 0 { [0, 1] } else { [1, 0] };
            for path in order {
                let mut output = vec![0; width * 8];
                let start = Instant::now();
                for _ in 0..4000 {
                    if path == 0 {
                        palette_scale_reference(black_box(&input), black_box(&mut output), width, 1, width,
                            width * 2, 1, width * 8, black_box(&colors)).unwrap();
                    } else {
                        double_palette_row(black_box(&input), black_box(&mut output), black_box(&colors)).unwrap();
                    }
                }
                let elapsed = start.elapsed().as_secs_f64();
                for (index, pair) in output.chunks_exact(8).enumerate() {
                    assert_eq!(&pair[..4], &colors[usize::from(input[index])]);
                    assert_eq!(&pair[4..], &colors[usize::from(input[index])]);
                }
                if round >= 2 { samples[path].push(elapsed); }
            }
        }
        for sample in &mut samples { sample.sort_by(f64::total_cmp); }
        eprintln!("palette-double width={width} generic_ms={:.3} double_ms={:.3} generic_over_double={:.3}",
            samples[0][8] * 1000.0, samples[1][8] * 1000.0, samples[0][8] / samples[1][8]);
    }
}
