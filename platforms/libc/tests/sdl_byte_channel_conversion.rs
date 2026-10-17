include!("../src/emitter/sdl_runtime.rs");

#[cfg(test)]
mod channel_conversion_tests {
    use super::*;

    fn formats() -> Vec<PixelFormat> {
        vec![
            PixelFormat { bytes: 3, masks: [0xff, 0xff00, 0xff0000, 0], palette: None },
            PixelFormat { bytes: 3, masks: [0xff0000, 0xff00, 0xff, 0], palette: None },
            PixelFormat { bytes: 4, masks: [0xff, 0xff00, 0xff0000, 0xff000000], palette: None },
            PixelFormat { bytes: 4, masks: [0xff0000, 0xff00, 0xff, 0xff000000], palette: None },
            PixelFormat { bytes: 4, masks: [0xff000000, 0xff0000, 0xff00, 0xff], palette: None },
        ]
    }

    fn object(fields: &[(&str, Value)]) -> Value {
        let mut value = Object::new();
        value.properties.extend(fields.iter().map(|(key, value)| ((*key).into(), value.clone())));
        Value::Object(heap::alloc(value))
    }

    fn surface(format: &PixelFormat, width: usize, rows: usize, pitch: usize, address: usize) -> Value {
        let format = object(&[("BitsPerPixel", Value::I32((format.bytes * 8) as i32)),
            ("Rmask", Value::I64(format.masks[0] as i64)),
            ("Gmask", Value::I64(format.masks[1] as i64)),
            ("Bmask", Value::I64(format.masks[2] as i64)),
            ("Amask", Value::I64(format.masks[3] as i64))]);
        object(&[("format", format), ("w", Value::I32(width as i32)),
            ("h", Value::I32(rows as i32)), ("pitch", Value::I32(pitch as i32)),
            ("pixels", Value::I32(address as i32))])
    }

    #[test]
    fn cropped_registered_conversions_match_snapshot_with_overlap_fallback() {
        use vybe_runtime::{Chunk, opcode::Op};
        let rectangle = |x, y| object(&[("x", Value::I32(x)), ("y", Value::I32(y)),
            ("w", Value::I32(5)), ("h", Value::I32(3))]);
        let formats = formats();
        for sf in &formats { for df in &formats {
            for (source_address, destination_address) in [(32usize, 700usize), (900, 0), (32, 0)] {
                let sp = 8 * sf.bytes + 3; let dp = 10 * df.bytes + 5;
                let seed: Vec<u8> = (0usize..4096).map(|i| (i * 37 + 129) as u8).collect();
                let mut expected = seed.clone();
                let reader = HostContext::empty();
                for y in 0..3 { for x in 0..5 {
                    let start = source_address + (y + 1) * sp + (x + 1) * sf.bytes;
                    let color = sf.decode(&reader, &seed, start).unwrap();
                    let start = destination_address + (y + 2) * dp + (x + 2) * df.bytes;
                    df.encode(color, &mut expected[start..start + df.bytes]).unwrap();
                }}
                let args = [surface(sf, 8, 6, sp, source_address), rectangle(1, 1),
                    surface(df, 10, 9, dp, destination_address), rectangle(2, 2)];
                let mut vm = VM::new(); vm.memory.with_buffer_mut(|memory| *memory = seed);
                register(&mut vm);
                let mut chunk = Chunk::new("cropped-channel-conversion");
                for (index, argument) in args.into_iter().enumerate() {
                    let name = format!("arg{index}");
                    vm.register_host_fn("test", &name, Box::new(move |_, _| argument.clone()));
                    let import = chunk.add_import("test", &name);
                    chunk.emit_call(import, 0, 0);
                }
                chunk.emit_i32_const(0, 0);
                let blit = chunk.add_import("libc:sdl", "convertBlit");
                chunk.emit_call(blit, 5, 0); chunk.emit_op(Op::RETURN, 0);
                assert_eq!(vm.run(vec![chunk]).unwrap(), Value::I32(0));
                assert_eq!(vm.memory.with_buffer(|memory| memory.to_vec()), expected,
                    "src={source_address} dst={destination_address} sf={:?} df={:?}", sf.masks, df.masks);
            }
        }}
    }

    #[test]
    fn byte_permutations_and_alpha_match_independent_mask_oracle() {
        let formats = formats();
        for sf in &formats { for df in &formats {
            for width in [1usize, 7, 256] { for rows in [1usize, 3] {
                let sp = width * sf.bytes + 3; let dp = width * df.bytes + 5;
                let span = rows * sp.max(dp);
                for reverse in [false, true] {
                    let (source, destination) = if reverse { (span + 64, 8) } else { (8, span + 64) };
                    let seed: Vec<u8> = (0usize..2 * span + 128)
                        .map(|i| (i.wrapping_mul(37) + 129) as u8).collect();
                    let mut actual = seed.clone(); let mut expected = seed.clone();
                    for y in 0..rows { for x in 0..width {
                        let mut bytes = [0u8; 4];
                        bytes[..sf.bytes].copy_from_slice(&seed[source + y * sp + x * sf.bytes
                            ..source + y * sp + (x + 1) * sf.bytes]);
                        let packed = u32::from_le_bytes(bytes);
                        let mut color = [0u8, 0, 0, 255];
                        for component in 0..4 {
                            if sf.masks[component] != 0 {
                                color[component] = ((packed & sf.masks[component])
                                    >> sf.masks[component].trailing_zeros()) as u8;
                            }
                        }
                        let mut result = 0u32;
                        for component in 0..4 {
                            if df.masks[component] != 0 {
                                result |= u32::from(color[component]) << df.masks[component].trailing_zeros();
                            }
                        }
                        let start = destination + y * dp + x * df.bytes;
                        expected[start..start + df.bytes].copy_from_slice(&result.to_le_bytes()[..df.bytes]);
                    }}
                    let mut ctx = HostContext::empty(); ctx.memory = Some(&mut actual);
                    assert_eq!(convert_linear_byte_rows(&mut ctx,
                        &Value::I32(source as i32), &Value::I32(destination as i32),
                        0, 0, width, rows, sp, dp, sf, df), Some(()));
                    assert_eq!(actual, expected);
                }
            }}
        }}
    }

    #[test]
    fn overlap_bad_pitch_and_invalid_span_leave_all_bytes_unchanged() {
        let formats = formats();
        for (source, destination, sp, dp) in [(8, 9, 24, 32), (8, 128, 20, 32),
            (8, 128, 24, 27), (250, 128, 24, 32)] {
            let mut memory = vec![0xa5; 256]; let before = memory.clone();
            let mut ctx = HostContext::empty(); ctx.memory = Some(&mut memory);
            assert_eq!(convert_linear_byte_rows(&mut ctx, &Value::I32(source), &Value::I32(destination),
                0, 0, 7, 3, sp, dp, &formats[0], &formats[2]), None);
            assert_eq!(memory, before);
        }
    }

    fn snapshot_reference(ctx: &mut HostContext<'_>, source: &Value, destination: &Value,
        width: usize, rows: usize, sp: usize, dp: usize, sf: &PixelFormat, df: &PixelFormat) -> Option<()> {
        let source = linear_memory32_address(source)? as usize;
        let destination = linear_memory32_address(destination)? as usize;
        let source_span = (rows - 1).checked_mul(sp)?.checked_add(width.checked_mul(sf.bytes)?)?;
        let destination_span = (rows - 1).checked_mul(dp)?.checked_add(width.checked_mul(df.bytes)?)?;
        let source_end = source.checked_add(source_span)?;
        let destination_end = destination.checked_add(destination_span)?;
        ctx.with_linear_memory_mut(|memory| {
            if source_end > memory.len() || destination_end > memory.len() { return None; }
            let input = memory[source..source_end].to_vec();
            let mut output = memory[destination..destination_end].to_vec();
            let reader = HostContext::empty();
            for y in 0..rows { for x in 0..width {
                let color = sf.decode(&reader, &input, y * sp + x * sf.bytes)?;
                let start = y * dp + x * df.bytes;
                df.encode(color, &mut output[start..start + df.bytes])?;
            }}
            memory[destination..destination_end].copy_from_slice(&output);
            Some(())
        })?
    }

    fn direct(ctx: &mut HostContext<'_>, source: &Value, destination: &Value,
        width: usize, rows: usize, sp: usize, dp: usize, sf: &PixelFormat, df: &PixelFormat) -> Option<()> {
        convert_linear_byte_rows(ctx, source, destination, 0, 0, width, rows, sp, dp, sf, df)
    }

    #[test]
    #[ignore = "optimized paired decode/encode benchmark, not Doom FPS"]
    fn paired_direct_conversion_against_existing_pixel_operations() {
        use std::{hint::black_box, time::Instant};
        assert!(!cfg!(debug_assertions));
        type Convert = fn(&mut HostContext<'_>, &Value, &Value, usize, usize, usize, usize,
            &PixelFormat, &PixelFormat) -> Option<()>;
        let formats = formats();
        for (source_format, destination_format) in [(0usize, 1usize), (0, 2), (2, 0), (2, 3), (2, 4)] {
            let sf = &formats[source_format]; let df = &formats[destination_format];
            let (width, rows) = (320usize, 200usize);
            let sp = width * sf.bytes + 3; let dp = width * df.bytes + 5;
            let source = Value::I32(8); let address = rows * sp + 64;
            let destination = Value::I32(address as i32);
            let seed: Vec<u8> = (0usize..address + rows * dp + 11)
                .map(|i| (i.wrapping_mul(37) + 129) as u8).collect();
            let mut reference = seed.clone(); let mut candidate = seed;
            let measure = |memory: &mut Vec<u8>, operation: Convert, iterations: usize| {
                let mut ctx = HostContext::empty(); ctx.memory = Some(memory);
                let start = Instant::now();
                for _ in 0..iterations {
                    black_box(operation)(&mut ctx, black_box(&source), black_box(&destination),
                        width, rows, sp, dp, black_box(sf), black_box(df)).unwrap();
                }
                start.elapsed().as_nanos() as f64
            };
            measure(&mut reference, snapshot_reference, 1);
            measure(&mut candidate, direct, 1);
            assert_eq!(candidate, reference);
            let mut ratios = Vec::new();
            for round in 0..19 {
                let (old, new) = if round % 2 == 0 {
                    (measure(&mut reference, snapshot_reference, 10), measure(&mut candidate, direct, 10))
                } else {
                    let new = measure(&mut candidate, direct, 10);
                    (measure(&mut reference, snapshot_reference, 10), new)
                };
                if round >= 2 { ratios.push(old / new); }
            }
            assert_eq!(candidate, reference);
            ratios.sort_by(f64::total_cmp);
            println!("320x200 channel conversion {source_format}->{destination_format}: {:.3}x",
                ratios[ratios.len()/2]);
        }
    }
}
