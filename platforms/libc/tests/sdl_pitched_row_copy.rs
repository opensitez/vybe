include!("../src/emitter/sdl_runtime.rs");

#[cfg(test)]
mod pitched_copy_tests {
    use super::*;

    fn object(fields: &[(&str, Value)]) -> Value {
        let mut value = Object::new();
        value.properties.extend(fields.iter().map(|(name, value)|
            ((*name).into(), value.clone())));
        Value::Object(heap::alloc(value))
    }

    #[test]
    fn convert_blit_rectangles_preserve_snapshot_results_and_padding() {
        let format = object(&[
            ("BitsPerPixel", Value::I32(32)),
            ("Rmask", Value::I64(255)), ("Gmask", Value::I64(65280)),
            ("Bmask", Value::I64(16711680)), ("Amask", Value::I64(4278190080)),
        ]);
        for destination in [0usize, 8, 16, 24, 32, 48, 64, 80] {
            for destination_pitch in [24usize, 28] {
                let source = object(&[("format", format.clone()), ("w", Value::I32(6)),
                    ("h", Value::I32(5)), ("pitch", Value::I32(24)), ("pixels", Value::I32(16))]);
                let target = object(&[("format", format.clone()), ("w", Value::I32(6)),
                    ("h", Value::I32(5)), ("pitch", Value::I32(destination_pitch as i32)),
                    ("pixels", Value::I32(destination as i32))]);
                let source_rect = object(&[("x", Value::I32(1)), ("y", Value::I32(1)),
                    ("w", Value::I32(3)), ("h", Value::I32(3))]);
                let target_rect = object(&[("x", Value::I32(0)), ("y", Value::I32(0)),
                    ("w", Value::I32(3)), ("h", Value::I32(3))]);
                let original: Vec<u8> = (0..256).map(|i| (i * 37 + 11) as u8).collect();
                let mut actual = original.clone();
                let mut expected = original.clone();
                for row in 0..3 {
                    let input = 16 + (row + 1) * 24 + 4;
                    let output = destination + row * destination_pitch;
                    expected[output..output + 12].copy_from_slice(&original[input..input + 12]);
                }
                let mut ctx = HostContext::empty();
                ctx.memory = Some(&mut actual);
                assert_eq!(convert_blit(&mut ctx,
                    &[source, source_rect, target, target_rect, Value::I32(0)]), Some(()));
                assert_eq!(actual, expected,
                    "destination={destination} pitch={destination_pitch}");
            }
        }
    }

    #[test]
    fn pitched_rows_match_snapshot_and_preserve_all_other_bytes() {
        for row_bytes in [4usize, 8, 12, 16] {
            for pitch in [16usize, 20, 24] {
                for rows in [1usize, 2, 3, 5] {
                    for destination in 0..=80 {
                        let source = 40usize;
                        let original: Vec<u8> = (0..256).map(|i| (i * 37 + 11) as u8).collect();
                        let mut actual = original.clone();
                        let mut expected = original.clone();
                        for row in 0..rows {
                            let start = source + row * pitch;
                            let target = destination + row * pitch;
                            expected[target..target + row_bytes]
                                .copy_from_slice(&original[start..start + row_bytes]);
                        }
                        let mut ctx = HostContext::empty();
                        ctx.memory = Some(&mut actual);
                        assert_eq!(copy_linear_rgba_rows(&mut ctx, &Value::I32(0),
                            &Value::I32(destination as i32), source, 0, row_bytes, rows, pitch), Some(()));
                        assert_eq!(actual, expected,
                            "bytes={row_bytes} pitch={pitch} rows={rows} dst={destination}");
                    }
                }
            }
        }
    }

    #[test]
    fn invalid_row_spans_do_not_modify_memory() {
        for (source, destination, row_bytes, rows, pitch) in [
            (0usize, 40usize, 16usize, 3usize, 24usize),
            (40, 0, 16, 3, 24), (0, 0, 25, 2, 24),
            (usize::MAX, 0, 4, 1, 4), (0, usize::MAX, 4, 1, 4),
            (0, 0, 4, usize::MAX, 24),
        ] {
            let mut memory = vec![0x5a; 64];
            let mut ctx = HostContext::empty();
            ctx.memory = Some(&mut memory);
            assert_eq!(copy_linear_rgba_rows(&mut ctx, &Value::I32(0),
                &Value::I32(0), source, destination, row_bytes, rows, pitch), None);
            assert_eq!(memory, vec![0x5a; 64]);
        }
    }
}
