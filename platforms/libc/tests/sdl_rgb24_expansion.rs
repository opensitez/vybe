include!("../src/emitter/sdl_runtime.rs");

fn expand_rgb24_rows(input: &[u8], width: usize, rows: usize, pitch: usize, bgr: bool) -> Option<Vec<u8>> {
    frame_kernels::expand_rgb24_rows(input, width, rows, pitch, bgr)
}

#[test]
fn rgb24_fast_path_requires_complete_rgb_masks_without_alpha_or_palette() {
    for (masks, expected) in [([0xff, 0xff00, 0xff0000, 0], Some(false)),
        ([0xff0000, 0xff00, 0xff, 0], Some(true)),
        ([0, 0xff00, 0xff0000, 0xff], None),
        ([0xff, 0xff00, 0xff0000, 0xff000000], None),
        ([0x7c00, 0x3e0, 0x1f, 0], None)] {
        let mut format = PixelFormat { bytes: 3, masks, palette: None };
        assert_eq!(format.rgb24_order(), expected);
        format.palette = Some(Value::Null);
        assert_eq!(format.rgb24_order(), None);
        format.palette = None;
        for bytes in [1, 2, 4] {
            format.bytes = bytes;
            assert_eq!(format.rgb24_order(), None);
        }
    }
}

fn mask_reference(input: &[u8], width: usize, rows: usize, pitch: usize, bgr: bool) -> Vec<u8> {
    let format = PixelFormat { bytes: 3, palette: None,
        masks: if bgr { [0xff0000, 0xff00, 0xff, 0] } else { [0xff, 0xff00, 0xff0000, 0] } };
    let ctx = HostContext::empty();
    let mut output = Vec::with_capacity(width * rows * 4);
    for row in 0..rows {
        for x in 0..width {
            output.extend_from_slice(&format.decode(&ctx, input, row * pitch + x * 3).unwrap());
        }
    }
    output
}

#[test]
fn rgb24_expansion_matches_mask_decoder_and_preserves_source() {
    for width in 0..18usize {
        for rows in 0..9usize {
            for padding in [0, 1, 13, 129] {
                let pitch = width * 3 + padding;
                let span = if rows == 0 { 0 } else { (rows - 1) * pitch + width * 3 };
                let input: Vec<u8> = (0..span).map(|index| (index * 37 + 91) as u8).collect();
                let before = input.clone();
                for bgr in [false, true] {
                    assert_eq!(expand_rgb24_rows(&input, width, rows, pitch, bgr),
                        Some(mask_reference(&input, width, rows, pitch, bgr)));
                    assert_eq!(input, before);
                }
            }
        }
    }
}

#[test]
fn rgb24_expansion_rejects_invalid_ranges_before_allocating_output() {
    let input = vec![0xa5; 23];
    for (width, rows, pitch) in [(4, 2, 11), (4, 2, 24), (usize::MAX, 1, usize::MAX),
        (1, usize::MAX, 3), (1, 3, usize::MAX)] {
        assert_eq!(expand_rgb24_rows(&input, width, rows, pitch, false), None);
        assert_eq!(input, vec![0xa5; 23]);
    }
}

#[test]
#[ignore = "paired production RGB24 expansion versus existing mask decoder; not Doom FPS"]
fn paired_rgb24_expansion_benchmark() {
    use std::hint::black_box;
    use std::time::Instant;
    for (width, rows, padding) in [(320usize, 200usize, 0usize), (320, 200, 13), (640, 480, 0), (640, 480, 13)] {
        for bgr in [false, true] {
            let pitch = width * 3 + padding;
            let span = (rows - 1) * pitch + width * 3;
            let input: Vec<u8> = (0..span).map(|index| (index * 17 + 39) as u8).collect();
            let expected = mask_reference(&input, width, rows, pitch, bgr);
            assert_eq!(expand_rgb24_rows(&input, width, rows, pitch, bgr), Some(expected));
            let mut samples = [Vec::new(), Vec::new()];
            for round in 0..19 {
                let order = if round % 2 == 0 { [0, 1] } else { [1, 0] };
                for path in order {
                    let start = Instant::now();
                    for _ in 0..20 {
                        let output = if path == 0 {
                            mask_reference(black_box(&input), width, rows, pitch, bgr)
                        } else {
                            expand_rgb24_rows(black_box(&input), width, rows, pitch, bgr).unwrap()
                        };
                        black_box(output);
                    }
                    let elapsed = start.elapsed().as_secs_f64();
                    if round >= 2 { samples[path].push(elapsed); }
                }
            }
            for sample in &mut samples { sample.sort_by(f64::total_cmp); }
            eprintln!("RGB24 {width}x{rows} padding={padding} bgr={bgr} mask_ms={:.3} expand_ms={:.3} mask_over_expand={:.3}",
                samples[0][8] * 1000.0, samples[1][8] * 1000.0, samples[0][8] / samples[1][8]);
        }
    }
}
