include!("../src/emitter/sdl_runtime.rs");

#[cfg(test)]
mod scaled_palette_tests {
    use super::*;

    fn palette() -> [[u8; 4]; 256] {
        std::array::from_fn(|i| [i as u8, (i * 37) as u8, (255 - i) as u8, (i * 13) as u8])
    }

    fn object(fields: &[(&str, Value)]) -> Value {
        let mut value = Object::new();
        value.properties.extend(fields.iter().map(|(key, value)| ((*key).into(), value.clone())));
        Value::Object(heap::alloc(value))
    }

    #[test]
    fn registered_scaled_palette_abi_matches_snapshot_in_both_orders_and_overlap() {
        use vybe_runtime::{Chunk, opcode::Op};
        let colors = palette();
        let mut entries = Object::new();
        entries.kind = ObjectKind::Array(colors.iter().map(|color| object(&[
            ("r", Value::I32(color[0] as i32)), ("g", Value::I32(color[1] as i32)),
            ("b", Value::I32(color[2] as i32)), ("a", Value::I32(color[3] as i32)),
        ])).collect());
        let palette = object(&[("colors", Value::Object(heap::alloc(entries))),
            ("ncolors", Value::I32(256)), ("__sdl_first", Value::I32(0))]);
        let indexed = object(&[("BitsPerPixel", Value::I32(8)), ("palette", palette),
            ("Rmask", Value::I32(0)), ("Gmask", Value::I32(0)),
            ("Bmask", Value::I32(0)), ("Amask", Value::I32(0))]);
        let rgba = object(&[("BitsPerPixel", Value::I32(32)),
            ("Rmask", Value::I64(255)), ("Gmask", Value::I64(65280)),
            ("Bmask", Value::I64(16711680)), ("Amask", Value::I64(4278190080))]);
        for (source_address, destination_address) in [(8usize, 256usize), (384, 8), (8, 9)] {
            let source = object(&[("format", indexed.clone()), ("w", Value::I32(4)),
                ("h", Value::I32(3)), ("pitch", Value::I32(7)),
                ("pixels", Value::I32(source_address as i32))]);
            let destination = object(&[("format", rgba.clone()), ("w", Value::I32(7)),
                ("h", Value::I32(5)), ("pitch", Value::I32(33)),
                ("pixels", Value::I32(destination_address as i32))]);
            let seed: Vec<u8> = (0usize..1024).map(|i| (i * 37 + 129) as u8).collect();
            let mut expected = seed.clone();
            for y in 0..5 { for x in 0..7 {
                let index = seed[source_address + (y * 3 / 5) * 7 + x * 4 / 7];
                let start = destination_address + y * 33 + x * 4;
                expected[start..start + 4].copy_from_slice(&colors[usize::from(index)]);
            }}
            let mut vm = VM::new(); vm.memory.with_buffer_mut(|memory| *memory = seed);
            register(&mut vm);
            vm.register_host_fn("test", "source", Box::new(move |_, _| source.clone()));
            vm.register_host_fn("test", "destination", Box::new(move |_, _| destination.clone()));
            let mut chunk = Chunk::new("scaled-palette-abi");
            let src = chunk.add_import("test", "source");
            let dst = chunk.add_import("test", "destination");
            let blit = chunk.add_import("libc:sdl", "convertBlit");
            chunk.emit_call(src, 0, 0); chunk.emit_ref_null(0x6f, 0);
            chunk.emit_call(dst, 0, 0); chunk.emit_ref_null(0x6f, 0);
            chunk.emit_i32_const(1, 0); chunk.emit_call(blit, 5, 0);
            chunk.emit_op(Op::RETURN, 0);
            assert_eq!(vm.run(vec![chunk]).unwrap(), Value::I32(0));
            assert_eq!(vm.memory.with_buffer(|memory| memory.to_vec()), expected);
        }
    }

    fn snapshot_reference(input: &[u8], output: &mut [u8], width: usize, rows: usize,
        pitch: usize, ow: usize, oh: usize, op: usize, colors: &[[u8; 4]]) -> Option<()> {
        // Fixed benchmark dimensions are valid. This intentionally omits the
        // converter's object lookup and per-pixel format encoding overhead.
        let input = input.to_vec();
        let mut pixels = output.to_vec();
        for y in 0..oh { for x in 0..ow {
            let index = input[(y * rows / oh) * pitch + x * width / ow];
            let start = y * op + x * 4;
            pixels[start..start + 4].copy_from_slice(&colors[usize::from(index)]);
        }}
        output.copy_from_slice(&pixels);
        Some(())
    }

    #[test]
    #[ignore = "optimized paired palette scaling benchmark, not Doom FPS"]
    fn paired_palette_scaling_against_snapshot_reference() {
        use std::{hint::black_box, time::Instant};
        assert!(!cfg!(debug_assertions));
        type Scale = fn(&[u8], &mut [u8], usize, usize, usize, usize, usize, usize, &[[u8; 4]]) -> Option<()>;
        let colors = palette();
        let (width, rows, pitch) = (320usize, 200usize, 323usize);
        let input: Vec<u8> = (0..rows * pitch).map(|i| (i * 37 + 129) as u8).collect();
        for (ow, oh) in [(640usize, 400usize), (960, 600), (1280, 800), (640, 100)] {
            let op = ow * 4 + 5;
            let mut reference = vec![0xa5; oh * op + 11];
            let mut candidate = reference.clone();
            snapshot_reference(&input, &mut reference, width, rows, pitch, ow, oh, op, &colors).unwrap();
            scale_palette_bytes(&input, &mut candidate, width, rows, pitch, ow, oh, op, &colors).unwrap();
            assert_eq!(reference, candidate);
            let mut ratios = Vec::new();
            for round in 0..19 {
                let measure = |output: &mut Vec<u8>, operation: Scale| {
                    let start = Instant::now();
                    for _ in 0..10 {
                        black_box(operation)(black_box(&input), black_box(output), width, rows,
                            pitch, ow, oh, op, black_box(&colors)).unwrap();
                    }
                    start.elapsed().as_nanos() as f64
                };
                let (old, new) = if round % 2 == 0 {
                    (measure(&mut reference, snapshot_reference), measure(&mut candidate, scale_palette_bytes))
                } else {
                    let new = measure(&mut candidate, scale_palette_bytes);
                    (measure(&mut reference, snapshot_reference), new)
                };
                if round >= 2 { ratios.push(old / new); }
            }
            assert_eq!(reference, candidate);
            ratios.sort_by(f64::total_cmp);
            println!("palette 320x200 -> {ow}x{oh}: {:.3}x", ratios[ratios.len()/2]);
        }
    }

    #[test]
    fn independent_mapping_oracle_preserves_alpha_and_padding() {
        let colors = palette();
        for width in 1usize..=17 { for output_width in 1usize..=19 {
            for rows in [1usize, 3, 7] { for output_rows in [1usize, 2, 9] {
                let pitch = width + 3; let output_pitch = output_width * 4 + 5;
                let input: Vec<u8> = (0..rows * pitch)
                    .map(|i| (i * 37 + 129) as u8).collect();
                let mut output = vec![0xa5; output_rows * output_pitch + 11];
                let mut expected = output.clone();
                for y in 0..output_rows { for x in 0..output_width {
                    let index = input[(y * rows / output_rows) * pitch + x * width / output_width];
                    let start = y * output_pitch + x * 4;
                    expected[start..start + 4].copy_from_slice(&colors[usize::from(index)]);
                }}
                assert_eq!(scale_palette_bytes(&input, &mut output, width, rows, pitch,
                    output_width, output_rows, output_pitch, &colors), Some(()));
                assert_eq!(output, expected);
            }}
        }}
    }

    #[test]
    fn short_palette_invalid_bounds_and_overlap_do_not_write() {
        let colors = palette();
        let input = vec![255; 24];
        for (width, rows, pitch, ow, oh, op, count) in [
            (4, 3, 8, 7, 5, 32, 255), (4, 3, 8, 7, 5, 27, 256),
            (4, 4, 8, 7, 5, 32, 256), (0, 3, 8, 7, 5, 32, 256),
            (usize::MAX, 3, 8, 7, 5, 32, 256),
        ] {
            let mut output = vec![0xa5; 160];
            assert_eq!(scale_palette_bytes(&input, &mut output, width, rows, pitch,
                ow, oh, op, &colors[..count]), None);
            assert!(output.iter().all(|&byte| byte == 0xa5));
        }
        let mut memory = vec![0xa5; 256];
        let mut ctx = HostContext::empty(); ctx.memory = Some(&mut memory);
        assert_eq!(scale_linear_palette_rows(&mut ctx, &Value::I32(8), &Value::I32(16),
            0, 0, 4, 3, 8, 7, 5, 32, &colors), None);
        assert!(memory.iter().all(|&byte| byte == 0xa5));
    }
}

#[test]
fn doom_argb_byte_channels_preserve_all_values() {
    let format = PixelFormat {
        bytes: 4,
        masks: [16711680, 65280, 255, 4278190080],
        palette: None,
    };
    assert!(format.byte_channels());
    for value in 0..=u8::MAX {
        let color = [value, value ^ 0x55, value.wrapping_mul(71), !value];
        let mut encoded = [0u8; 4];
        format.encode(color, &mut encoded);
        assert_eq!(encoded, [color[2], color[1], color[0], color[3]]);
    }
}
