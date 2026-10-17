include!("../src/emitter/sdl_runtime.rs");

#[cfg(test)]
mod matching_byte_rows {
    use super::*;

    fn object(fields: &[(&str, Value)]) -> Value {
        let mut object = Object::new();
        object.properties.extend(fields.iter().map(|(k, v)| ((*k).into(), v.clone())));
        Value::Object(heap::alloc(object))
    }

    fn surface(bytes: usize, bgr: bool, width: usize, rows: usize, pitch: usize,
        address: usize) -> Value {
        let (red, blue) = if bgr { (0xff0000, 0xff) } else { (0xff, 0xff0000) };
        let format = object(&[
            ("BitsPerPixel", Value::I32((bytes * 8) as i32)),
            ("Rmask", Value::I64(red)), ("Gmask", Value::I64(0xff00)),
            ("Bmask", Value::I64(blue)),
            ("Amask", Value::I64(if bytes == 4 { 0xff000000 } else { 0 })),
        ]);
        object(&[("format", format), ("w", Value::I32(width as i32)),
            ("h", Value::I32(rows as i32)), ("pitch", Value::I32(pitch as i32)),
            ("pixels", Value::I32(address as i32))])
    }

    #[test]
    fn identical_byte_formats_match_snapshot_including_cross_row_overlap() {
        for bytes in [3usize, 4] { for bgr in [false, true] {
            for width in [1usize, 3, 7] { for rows in [1usize, 3, 5] {
                for padding in [0usize, 1, 9] {
                    let row_bytes = width * bytes;
                    let pitch = row_bytes + padding;
                    let source = 192usize;
                    for displacement in [-(pitch as isize), -1, 0, 1, pitch as isize, 256] {
                        let destination = source.checked_add_signed(displacement).unwrap();
                        let seed: Vec<u8> = (0usize..1024)
                            .map(|i| (i.wrapping_mul(37) + 129) as u8).collect();
                        let mut actual = seed.clone();
                        let mut expected = seed.clone();
                        for row in 0..rows {
                            expected[destination + row * pitch..destination + row * pitch + row_bytes]
                                .copy_from_slice(&seed[source + row * pitch..source + row * pitch + row_bytes]);
                        }
                        let args = [surface(bytes, bgr, width, rows, pitch, source), Value::Null,
                            surface(bytes, bgr, width, rows, pitch, destination), Value::Null,
                            Value::I32(0)];
                        let mut ctx = HostContext::empty();
                        ctx.memory = Some(&mut actual);
                        assert_eq!(convert_blit(&mut ctx, &args), Some(()));
                        assert_eq!(actual, expected,
                            "bytes={bytes} bgr={bgr} width={width} rows={rows} padding={padding} displacement={displacement}");
                    }
                }
            }}
        }}
    }

    #[test]
    fn unequal_pitches_preserve_padding_and_overlapping_snapshot_fallback() {
        for bytes in [3usize, 4] { for bgr in [false, true] {
            for width in [1usize, 3, 7] { for rows in [1usize, 3, 5] {
                for source_padding in [0usize, 7] { for destination_padding in [0usize, 11] {
                    let row_bytes = width * bytes;
                    let sp = row_bytes + source_padding;
                    let dp = row_bytes + destination_padding;
                    for (source, destination) in [(192usize, 191usize), (192, 192),
                        (192, 193), (192, 448), (448, 192)] {
                        let seed: Vec<u8> = (0usize..1024)
                            .map(|i| (i.wrapping_mul(37) + 129) as u8).collect();
                        let mut actual = seed.clone();
                        let mut expected = seed.clone();
                        for row in 0..rows {
                            expected[destination + row * dp..destination + row * dp + row_bytes]
                                .copy_from_slice(&seed[source + row * sp..source + row * sp + row_bytes]);
                        }
                        let args = [surface(bytes, bgr, width, rows, sp, source), Value::Null,
                            surface(bytes, bgr, width, rows, dp, destination), Value::Null,
                            Value::I32(0)];
                        let mut ctx = HostContext::empty(); ctx.memory = Some(&mut actual);
                        assert_eq!(convert_blit(&mut ctx, &args), Some(()));
                        assert_eq!(actual, expected,
                            "bytes={bytes} bgr={bgr} width={width} rows={rows} sp={sp} dp={dp} src={source} dst={destination}");
                    }
                }}
            }}
        }}
    }

    #[test]
    fn different_channel_order_still_converts_instead_of_raw_copy() {
        for bytes in [3usize, 4] {
            let mut memory = vec![0xa5; 128];
            memory[8..8 + bytes].copy_from_slice(&[17, 53, 211, 127][..bytes]);
            let args = [surface(bytes, false, 1, 1, bytes, 8), Value::Null,
                surface(bytes, true, 1, 1, bytes, 64), Value::Null, Value::I32(0)];
            let mut ctx = HostContext::empty();
            ctx.memory = Some(&mut memory);
            assert_eq!(convert_blit(&mut ctx, &args), Some(()));
            assert_eq!(&memory[64..64 + bytes], &[211, 53, 17, 127][..bytes]);
        }
    }
}
