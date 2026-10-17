include!("../src/emitter/sdl_runtime.rs");

fn single_pass_frame(input: &[u8], width: usize, rows: usize, pitch: usize) -> Option<Vec<u8>> {
    let row_bytes = width.checked_mul(4)?;
    let length = row_bytes.checked_mul(rows)?;
    if rows != 0 && pitch < row_bytes { return None; }
    let span = if rows == 0 { 0 } else {
        (rows - 1).checked_mul(pitch)?.checked_add(row_bytes)?
    };
    if span > input.len() { return None; }
    let mut output = Vec::new();
    output.try_reserve_exact(length).ok()?;
    if row_bytes != 0 {
        for row in 0..rows {
            let start = row * pitch;
            let source = &input[start..start + row_bytes];
            let mut words = source.chunks_exact(8);
            output.extend(words.by_ref().flat_map(|word| {
                let mut bytes: [u8; 8] = word.try_into().unwrap();
                frame_kernels::swap_bgra_bytes(&mut bytes);
                bytes
            }));
            for pixel in words.remainder().chunks_exact(4) {
                output.extend_from_slice(&[pixel[2], pixel[1], pixel[0], pixel[3]]);
            }
        }
    }
    Some(output)
}

#[test]
fn single_pass_matches_existing_checked_copy_and_preserves_input() {
    for width in 0..18usize {
        for rows in 0..9usize {
            for padding in [0, 1, 13, 129] {
                let pitch = width * 4 + padding;
                let span = if rows == 0 { 0 } else { (rows - 1) * pitch + width * 4 };
                let input: Vec<u8> = (0..span).map(|index| (index * 37 + 91) as u8).collect();
                let before = input.clone();
                assert_eq!(single_pass_frame(&input, width, rows, pitch),
                    compact_rgba_rows(&input, width, rows, pitch, true));
                assert_eq!(input, before);
            }
        }
    }
    let input = vec![0xa5; 31];
    for (width, rows, pitch) in [(4, 2, 15), (4, 2, 32), (usize::MAX, 1, usize::MAX),
        (1, usize::MAX, 4), (1, 3, usize::MAX)] {
        assert_eq!(single_pass_frame(&input, width, rows, pitch), None);
        assert_eq!(input, vec![0xa5; 31]);
    }
}

#[test]
#[ignore = "pure one-pass conversion experiment; does not establish a VM borrowing API or Doom speedup"]
fn paired_single_pass_conversion_benchmark() {
    use std::hint::black_box;
    use std::time::Instant;
    for (width, rows, padding) in [(320usize, 200usize, 0usize), (320, 200, 13), (640, 480, 0), (640, 480, 13)] {
        let pitch = width * 4 + padding;
        let span = (rows - 1) * pitch + width * 4;
        let input: Vec<u8> = (0..span).map(|index| (index * 17 + 39) as u8).collect();
        assert_eq!(single_pass_frame(&input, width, rows, pitch),
            compact_rgba_rows(&input, width, rows, pitch, true));
        let mut samples = [Vec::new(), Vec::new()];
        for round in 0..19 {
            let order = if round % 2 == 0 { [0, 1] } else { [1, 0] };
            for path in order {
                let start = Instant::now();
                for _ in 0..40 {
                    let output = if path == 0 {
                        compact_rgba_rows(black_box(&input), width, rows, pitch, true)
                    } else {
                        single_pass_frame(black_box(&input), width, rows, pitch)
                    }.unwrap();
                    black_box(output);
                }
                let elapsed = start.elapsed().as_secs_f64();
                if round >= 2 { samples[path].push(elapsed); }
            }
        }
        for sample in &mut samples { sample.sort_by(f64::total_cmp); }
        eprintln!("single-pass {width}x{rows} padding={padding} bulk_ms={:.3} single_ms={:.3} bulk_over_single={:.3}",
            samples[0][8] * 1000.0, samples[1][8] * 1000.0, samples[0][8] / samples[1][8]);
    }
}
