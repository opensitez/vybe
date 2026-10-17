include!("../src/emitter/sdl_runtime.rs");

fn word_swap(bytes: &mut [u8]) {
    frame_kernels::swap_bgra_bytes(bytes);
}

fn pixel_pack(bytes: &mut Vec<u8>, width: usize, rows: usize, pitch: usize) -> Option<()> {
    pack_rgba_rows_in_place(bytes, width, rows, pitch, false)?;
    for pixel in bytes.chunks_exact_mut(4) {
        pixel.swap(0, 2);
    }
    Some(())
}

fn candidate_pack(bytes: &mut Vec<u8>, width: usize, rows: usize, pitch: usize) -> Option<()> {
    pack_rgba_rows_in_place(bytes, width, rows, pitch, false)?;
    word_swap(bytes);
    Some(())
}

#[test]
fn word_swap_matches_independent_pixels_and_preserves_partial_tail() {
    for length in 0..=71 {
        for seed in 0..=255u16 {
            let before: Vec<u8> = (0..length).map(|index| {
                (index * 73 + usize::from(seed)) as u8
            }).collect();
            let mut expected = before.clone();
            for start in (0..length / 4).map(|pixel| pixel * 4) {
                expected[start] = before[start + 2];
                expected[start + 2] = before[start];
            }
            let mut actual = before.clone();
            word_swap(&mut actual);
            assert_eq!(actual, expected, "length={length} seed={seed}");
            word_swap(&mut actual);
            assert_eq!(actual, before);
        }
    }
}

#[test]
fn word_packing_matches_existing_checked_row_packing() {
    for width in 0..18usize {
        for rows in 0..9usize {
            for padding in [0, 1, 3, 13] {
                let pitch = width * 4 + padding;
                let span = if rows == 0 { 0 } else { (rows - 1) * pitch + width * 4 };
                let before: Vec<u8> = (0..span).map(|index| (index * 37 + 91) as u8).collect();
                let mut expected = before.clone();
                let mut actual = before;
                let pointer = actual.as_ptr();
                let capacity = actual.capacity();
                assert_eq!(pixel_pack(&mut expected, width, rows, pitch), Some(()));
                assert_eq!(candidate_pack(&mut actual, width, rows, pitch), Some(()));
                assert_eq!(actual, expected);
                assert_eq!(actual.as_ptr(), pointer);
                assert_eq!(actual.capacity(), capacity);
            }
        }
    }
}

#[test]
#[ignore = "paired production BGRA kernel and owned packing versus per-pixel reference; not Doom FPS"]
fn paired_word_swap_and_owned_frame_benchmark() {
    use std::hint::black_box;
    use std::time::Instant;

    for (width, rows, padding) in [(320usize, 200usize, 0usize), (320, 200, 13), (640, 480, 0), (640, 480, 13)] {
        let pitch = width * 4 + padding;
        let span = (rows - 1) * pitch + width * 4;
        let source: Vec<u8> = (0..span).map(|index| (index * 17 + 39) as u8).collect();
        for kernel_only in [true, false] {
            let mut samples = [Vec::new(), Vec::new()];
            for round in 0..19 {
                let order = if round % 2 == 0 { [0, 1] } else { [1, 0] };
                for path in order {
                    let mut pixels = source.clone();
                    let start = Instant::now();
                    for _ in 0..40 {
                        if kernel_only {
                            if path == 0 {
                                for pixel in black_box(pixels.as_mut_slice()).chunks_exact_mut(4) {
                                    pixel.swap(0, 2);
                                }
                            } else {
                                word_swap(black_box(pixels.as_mut_slice()));
                            }
                        } else {
                            let mut snapshot = black_box(&source).clone();
                            if path == 0 {
                                pixel_pack(&mut snapshot, width, rows, pitch).unwrap();
                            } else {
                                candidate_pack(&mut snapshot, width, rows, pitch).unwrap();
                            }
                            black_box(snapshot);
                        }
                    }
                    let elapsed = start.elapsed().as_secs_f64();
                    if kernel_only { assert_eq!(pixels, source); }
                    if round >= 2 { samples[path].push(elapsed); }
                }
            }
            for sample in &mut samples { sample.sort_by(f64::total_cmp); }
            eprintln!("BGRA-word {width}x{rows} padding={padding} kernel_only={kernel_only} pixel_ms={:.3} word_ms={:.3} pixel_over_word={:.3}",
                samples[0][8] * 1000.0, samples[1][8] * 1000.0, samples[0][8] / samples[1][8]);
        }
    }
}
