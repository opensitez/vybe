include!("../src/emitter/sdl_runtime.rs");

#[test]
fn compact_copy_matches_independent_pixels_and_preserves_source() {
    for width in 0..18usize {
        for rows in 0..9usize {
            for padding in [0, 1, 13, width * 4 + 17] {
                let pitch = width * 4 + padding;
                let span = if rows == 0 { 0 } else { (rows - 1) * pitch + width * 4 };
                let source: Vec<u8> = (0..span).map(|index| (index * 37 + 91) as u8).collect();
                let before = source.clone();
                for bgra in [false, true] {
                    let mut expected = Vec::new();
                    for row in 0..rows {
                        for x in 0..width {
                            let start = row * pitch + x * 4;
                            let pixel = &source[start..start + 4];
                            expected.extend_from_slice(if bgra {
                                &[pixel[2], pixel[1], pixel[0], pixel[3]]
                            } else {
                                pixel
                            });
                        }
                    }
                    assert_eq!(compact_rgba_rows(&source, width, rows, pitch, bgra), Some(expected));
                    assert_eq!(source, before);
                }
            }
        }
    }
}

#[test]
fn compact_copy_rejects_invalid_ranges_without_changing_source() {
    let source = vec![0xa5; 31];
    for (width, rows, pitch) in [
        (4, 2, 15), (4, 2, 32), (usize::MAX, 1, usize::MAX),
        (1, usize::MAX, 4), (1, 3, usize::MAX),
    ] {
        assert_eq!(compact_rgba_rows(&source, width, rows, pitch, true), None);
        assert_eq!(source, vec![0xa5; 31]);
    }
}

#[test]
fn oversized_source_padding_is_not_retained_by_compact_output() {
    let source = vec![0x83; 4096 + 8];
    let output = compact_rgba_rows(&source, 2, 2, 4096, true).unwrap();
    assert_eq!(output, vec![0x83; 16]);
    assert!(!reuse_owned_frame_buffer(source.capacity(), output.len()));
    assert!(output.capacity() < source.capacity());
}
