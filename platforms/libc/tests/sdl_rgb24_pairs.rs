include!("../src/emitter/sdl_runtime.rs");

fn expand_pairs(input: &[u8], width: usize, rows: usize, pitch: usize, bgr: bool) -> Option<Vec<u8>> {
    let row_bytes = width.checked_mul(3)?;
    let length = width.checked_mul(rows)?.checked_mul(4)?;
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
            let mut pairs = input[start..start + row_bytes].chunks_exact(6);
            if bgr {
                output.extend(pairs.by_ref().flat_map(|p| [p[2], p[1], p[0], 255, p[5], p[4], p[3], 255]));
            } else {
                output.extend(pairs.by_ref().flat_map(|p| [p[0], p[1], p[2], 255, p[3], p[4], p[5], 255]));
            }
            for p in pairs.remainder().chunks_exact(3) {
                let red = p[if bgr { 2 } else { 0 }];
                let blue = p[if bgr { 0 } else { 2 }];
                output.extend_from_slice(&[red, p[1], blue, 255]);
            }
        }
    }
    Some(output)
}

#[test]
fn paired_rgb24_expansion_matches_production_and_preserves_source() {
    for width in 0..34usize {
        for rows in 0..9usize {
            for padding in [0, 1, 13, 129] {
                let pitch = width * 3 + padding;
                let span = if rows == 0 { 0 } else { (rows - 1) * pitch + width * 3 };
                let input: Vec<u8> = (0..span).map(|index| (index * 37 + 91) as u8).collect();
                let before = input.clone();
                for bgr in [false, true] {
                    assert_eq!(expand_pairs(&input, width, rows, pitch, bgr),
                        frame_kernels::expand_rgb24_rows(&input, width, rows, pitch, bgr));
                    assert_eq!(input, before);
                }
            }
        }
    }
    let input = [0xa5; 23];
    for (width, rows, pitch) in [(4, 2, 11), (4, 2, 24), (usize::MAX, 1, usize::MAX),
        (1, usize::MAX, 3), (1, 3, usize::MAX)] {
        assert_eq!(expand_pairs(&input, width, rows, pitch, false), None);
        assert_eq!(input, [0xa5; 23]);
    }
}

#[test]
#[ignore = "paired two-pixel RGB24 experiment versus production expansion; not Doom timing"]
fn paired_rgb24_pair_kernel_benchmark() {
    use std::hint::black_box;
    use std::time::Instant;
    for (width, rows, padding) in [(320usize, 200usize, 0usize), (321, 200, 13), (640, 480, 0), (641, 480, 13)] {
        for bgr in [false, true] {
            let pitch = width * 3 + padding;
            let span = (rows - 1) * pitch + width * 3;
            let input: Vec<u8> = (0..span).map(|index| (index * 17 + 39) as u8).collect();
            let mut samples = [Vec::new(), Vec::new()];
            for round in 0..19 {
                let order = if round % 2 == 0 { [0, 1] } else { [1, 0] };
                for path in order {
                    let start = Instant::now();
                    for _ in 0..40 {
                        let output = if path == 0 {
                            frame_kernels::expand_rgb24_rows(black_box(&input), width, rows, pitch, bgr)
                        } else {
                            expand_pairs(black_box(&input), width, rows, pitch, bgr)
                        }.unwrap();
                        black_box(output);
                    }
                    let elapsed = start.elapsed().as_secs_f64();
                    if round >= 2 { samples[path].push(elapsed); }
                }
            }
            for sample in &mut samples { sample.sort_by(f64::total_cmp); }
            eprintln!("RGB24-pair {width}x{rows} padding={padding} bgr={bgr} production_ms={:.3} pair_ms={:.3} production_over_pair={:.3}",
                samples[0][8] * 1000.0, samples[1][8] * 1000.0, samples[0][8] / samples[1][8]);
        }
    }
}
