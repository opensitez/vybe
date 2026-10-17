include!("../src/emitter/sdl_runtime.rs");

fn packed_reference(input: &[u8], width: usize, rows: usize, pitch: usize, bgra: bool) -> Vec<u8> {
    let mut output = Vec::with_capacity(width * rows * 4);
    for row in 0..rows {
        let source = &input[row * pitch..row * pitch + width * 4];
        if bgra {
            for pixel in source.chunks_exact(4) {
                output.extend_from_slice(&[pixel[2], pixel[1], pixel[0], pixel[3]]);
            }
        } else {
            output.extend_from_slice(source);
        }
    }
    output
}

#[test]
fn owned_frame_packing_matches_rows_without_reallocation() {
    for width in 0..18usize {
        for rows in 0..9usize {
            for padding in [0, 1, 3, 13] {
                let pitch = width * 4 + padding;
                let span = if rows == 0 { 0 } else { (rows - 1) * pitch + width * 4 };
                let input: Vec<u8> = (0..span).map(|index| (index * 37 + 91) as u8).collect();
                for bgra in [false, true] {
                    let expected = packed_reference(&input, width, rows, pitch, bgra);
                    let mut actual = input.clone();
                    let pointer = actual.as_ptr();
                    let capacity = actual.capacity();
                    assert_eq!(pack_rgba_rows_in_place(&mut actual, width, rows, pitch, bgra), Some(()));
                    assert_eq!(actual, expected);
                    assert_eq!(actual.as_ptr(), pointer);
                    assert_eq!(actual.capacity(), capacity);
                }
            }
        }
    }
}

#[test]
fn invalid_owned_frame_ranges_do_not_mutate_snapshot() {
    for (width, rows, pitch) in [
        (4, 2, 15), (4, 2, 32), (usize::MAX, 1, usize::MAX),
        (1, usize::MAX, 4), (1, 3, usize::MAX),
    ] {
        let mut bytes = vec![0xa5; 31];
        let before = bytes.clone();
        let capacity = bytes.capacity();
        assert_eq!(pack_rgba_rows_in_place(&mut bytes, width, rows, pitch, true), None);
        assert_eq!(bytes, before);
        assert_eq!(bytes.capacity(), capacity);
    }
}

#[test]
fn owned_buffer_reuse_bounds_retained_padding_capacity() {
    assert!(reuse_owned_frame_buffer(100, 100));
    assert!(reuse_owned_frame_buffer(112, 100));
    assert!(!reuse_owned_frame_buffer(113, 100));
    assert!(!reuse_owned_frame_buffer(4096, 8));
    assert!(!reuse_owned_frame_buffer(0, 0));
    assert!(!reuse_owned_frame_buffer(4096, 0));
    assert!(!reuse_owned_frame_buffer(usize::MAX, usize::MAX));
}

#[test]
#[ignore = "paired owned framebuffer conversion; not whole SDL presentation or Doom FPS"]
fn paired_owned_frame_packing_benchmark() {
    use std::hint::black_box;
    use std::time::Instant;
    for (width, rows) in [(320usize, 200usize), (640, 480)] {
        for (bgra, padding) in [(false, 13usize), (true, 0), (true, 13), (true, width)] {
            let pitch = width * 4 + padding;
            let span = (rows - 1) * pitch + width * 4;
            let source: Vec<u8> = (0..span).map(|index| (index * 17 + 39) as u8).collect();
            let expected = packed_reference(&source, width, rows, pitch, bgra);
            let mut probe = source.clone();
            pack_rgba_rows_in_place(&mut probe, width, rows, pitch, bgra).unwrap();
            assert_eq!(probe, expected);
            let mut samples = [Vec::new(), Vec::new()];
            for round in 0..19 {
                let order = if round % 2 == 0 { [0, 1] } else { [1, 0] };
                for path in order {
                    let start = Instant::now();
                    for _ in 0..40 {
                        let mut snapshot = black_box(&source).clone();
                        let output = if path == 0 {
                            packed_reference(&snapshot, width, rows, pitch, bgra)
                        } else if !bgra
                            || !reuse_owned_frame_buffer(snapshot.capacity(), width * rows * 4) {
                            compact_rgba_rows(&snapshot, width, rows, pitch, bgra).unwrap()
                        } else {
                            pack_rgba_rows_in_place(&mut snapshot, width, rows, pitch, bgra).unwrap();
                            snapshot
                        };
                        black_box(output);
                    }
                    let elapsed = start.elapsed().as_secs_f64();
                    if round >= 2 { samples[path].push(elapsed); }
                }
            }
            for sample in &mut samples { sample.sort_by(f64::total_cmp); }
            eprintln!("owned-frame {width}x{rows} bgra={bgra} padding={padding} snapshot_ms={:.3} inplace_ms={:.3} speedup={:.3}",
                samples[0][8] * 1000.0, samples[1][8] * 1000.0, samples[0][8] / samples[1][8]);
        }
    }
}
