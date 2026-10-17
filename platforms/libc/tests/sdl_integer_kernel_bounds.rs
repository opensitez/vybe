#[path = "../src/emitter/sdl_pixel_kernels.rs"]
mod kernels;

fn check_mapping<const FACTOR: usize>() {
    for width in 1..=9usize {
        for rows in 1..=7usize {
            for output_rows in 1..=13usize {
                let pitch = width * 4 + 3;
                let source_end = (rows - 1) * pitch + width * 4;
                let source: Vec<u8> = (0..source_end).map(|i| (i * 73 + 19) as u8).collect();
                let original_source = source.clone();
                let output_width = width * FACTOR;
                let output_pitch = output_width * 4 + 5;
                let output_end = (output_rows - 1) * output_pitch + output_width * 4;
                let mut output = vec![0x91; output_end + 17];
                let mut expected = output.clone();
                for y in 0..output_rows {
                    let source_y = y * rows / output_rows;
                    for x in 0..output_width {
                        let source_start = source_y * pitch + (x / FACTOR) * 4;
                        let destination_start = y * output_pitch + x * 4;
                        expected[destination_start..destination_start + 4]
                            .copy_from_slice(&source[source_start..source_start + 4]);
                    }
                }
                assert_eq!(kernels::scale_integer_rgba_bytes::<FACTOR>(
                    &source, &mut output, width, rows, pitch, output_rows, output_pitch,
                ), Some(()));
                assert_eq!(output, expected);
                assert_eq!(source, original_source);
            }
        }
    }
}

#[test]
fn both_integer_factors_match_independent_mapping_and_guards() {
    check_mapping::<3>();
    check_mapping::<4>();
}

fn check_rejections<const FACTOR: usize>() {
    let source = [1u8; 64];
    for (width, rows, pitch, output_rows, output_pitch) in [
        (0, 1, 4, 1, 16), (1, 0, 4, 1, 16), (1, 1, 4, 0, 16),
        (1, 1, 3, 1, 16), (1, 1, 4, 1, 11), (2, 9, 8, 1, 32),
        (2, 1, 8, 6, 32), (usize::MAX, 1, 4, 1, 16),
        (1, usize::MAX, 4, 1, 16), (1, 1, 4, usize::MAX, 16),
        (1, 2, usize::MAX, 1, 16), (1, 1, 4, 2, usize::MAX),
    ] {
        let mut output = [0x91; 128];
        assert_eq!(kernels::scale_integer_rgba_bytes::<FACTOR>(
            &source, &mut output, width, rows, pitch, output_rows, output_pitch,
        ), None);
        assert_eq!(output, [0x91; 128]);
    }
    let mut output = [0x91; 128];
    assert_eq!(kernels::scale_integer_rgba_bytes::<FACTOR>(
        &source[..3], &mut output, 1, 1, 4, 1, 16,
    ), None);
    assert_eq!(output, [0x91; 128]);
    assert_eq!(kernels::scale_integer_rgba_bytes::<FACTOR>(
        &source, &mut output[..4 * FACTOR - 1], 1, 1, 4, 1, 4 * FACTOR,
    ), None);
    assert_eq!(output, [0x91; 128]);
}

#[test]
fn invalid_geometry_and_ranges_reject_before_writes() {
    check_rejections::<3>();
    check_rejections::<4>();
}

#[test]
fn unsupported_factors_reject_before_writes() {
    let mut output = [0x91; 32];
    for result in [
        kernels::scale_integer_rgba_bytes::<0>(&[1; 4], &mut output, 1, 1, 4, 1, 32),
        kernels::scale_integer_rgba_bytes::<1>(&[1; 4], &mut output, 1, 1, 4, 1, 32),
        kernels::scale_integer_rgba_bytes::<2>(&[1; 4], &mut output, 1, 1, 4, 1, 32),
        kernels::scale_integer_rgba_bytes::<5>(&[1; 4], &mut output, 1, 1, 4, 1, 32),
    ] {
        assert_eq!(result, None);
    }
    assert_eq!(output, [0x91; 32]);
}
