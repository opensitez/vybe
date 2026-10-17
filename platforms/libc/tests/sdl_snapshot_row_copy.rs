#[path = "../src/emitter/sdl_pixel_kernels.rs"]
mod kernels;

#[test]
fn packed_and_pitched_snapshots_preserve_all_guards_and_source() {
    for bytes in [3, 4] {
        for width in 1..=11 {
            for rows in 1..=9 {
                for source_padding in [0, 1, 9] {
                    for destination_padding in [0, 3, 13] {
                        let row_bytes = width * bytes;
                        let source_pitch = row_bytes + source_padding;
                        let destination_pitch = row_bytes + destination_padding;
                        let source_span = (rows - 1) * source_pitch + row_bytes;
                        let destination_span = (rows - 1) * destination_pitch + row_bytes;
                        let source: Vec<u8> = (0..source_span + 9)
                            .map(|i| (i * 73 + 19) as u8).collect();
                        let source_before = source.clone();
                        let mut output = vec![0x91; destination_span + 17];
                        let mut expected = output.clone();
                        for row in 0..rows {
                            for byte in 0..row_bytes {
                                expected[row * destination_pitch + byte] = source[row * source_pitch + byte];
                            }
                        }
                        assert_eq!(kernels::copy_snapshot_byte_rows(
                            &source[..source_span], &mut output, row_bytes,
                            rows, source_pitch, destination_pitch,
                        ), Some(()));
                        assert_eq!(output, expected);
                        assert_eq!(source, source_before);
                    }
                }
            }
        }
    }
}

#[test]
fn invalid_snapshot_spans_reject_before_any_write() {
    let source = [1u8; 64];
    for (row_bytes, rows, source_pitch, destination_pitch) in [
        (4, 1, 3, 4), (4, 1, 4, 3), (4, 17, 4, 4), (4, 33, 4, 4),
        (usize::MAX, 1, usize::MAX, usize::MAX),
        (4, usize::MAX, 4, 4), (4, 2, usize::MAX, 4), (4, 2, 4, usize::MAX),
    ] {
        let mut output = [0x91; 128];
        assert_eq!(kernels::copy_snapshot_byte_rows(
            &source, &mut output, row_bytes, rows, source_pitch, destination_pitch,
        ), None);
        assert_eq!(output, [0x91; 128]);
    }
    let mut short_output = [0x91; 3];
    assert_eq!(kernels::copy_snapshot_byte_rows(&source, &mut short_output, 4, 1, 4, 4), None);
    assert_eq!(short_output, [0x91; 3]);
}

#[test]
fn zero_extents_are_noops() {
    let mut output = [0x91; 8];
    assert_eq!(kernels::copy_snapshot_byte_rows(&[], &mut output, 0, usize::MAX, 0, 0), Some(()));
    assert_eq!(kernels::copy_snapshot_byte_rows(&[], &mut output, usize::MAX, 0, 0, 0), Some(()));
    assert_eq!(output, [0x91; 8]);
}
