include!("../src/emitter/sdl_runtime.rs");

#[cfg(test)]
mod scaled_rgba_tests {
    use super::*;
    use vybe_runtime::{Chunk, opcode::Op};

    #[test]
    fn rational_mapping_matches_independent_oracle_and_preserves_padding() {
        for width in 1..=17 {
            for output_width in 1..=19 {
                for rows in [1, 3, 7] {
                    for output_rows in [1, 2, 9] {
                        let pitch = width * 4 + 3;
                        let output_pitch = output_width * 4 + 5;
                        let input: Vec<u8> = (0..rows * pitch + 11)
                            .map(|i: usize| (i.wrapping_mul(37) + 11) as u8).collect();
                        let mut output = vec![0xa5; output_rows * output_pitch + 17];
                        let mut expected = output.clone();
                        for y in 0..output_rows {
                            for x in 0..output_width {
                                let source = (y * rows / output_rows) * pitch
                                    + (x * width / output_width) * 4;
                                let destination = y * output_pitch + x * 4;
                                expected[destination..destination + 4]
                                    .copy_from_slice(&input[source..source + 4]);
                            }
                        }
                        assert_eq!(scale_rgba_bytes(&input, &mut output, width, rows, pitch,
                            output_width, output_rows, output_pitch), Some(()));
                        assert_eq!(output, expected);
                    }
                }
            }
        }
    }

    #[test]
    fn invalid_spans_and_dimensions_leave_all_output_unchanged() {
        for (input_len, output_len, width, rows, pitch, ow, oh, op) in [
            (55, 136, 4, 3, 20, 6, 5, 28),
            (56, 135, 4, 3, 20, 6, 5, 28),
            (56, 136, 4, 3, 15, 6, 5, 28),
            (56, 136, 4, 3, 20, 6, 5, 23),
            (56, 136, 0, 3, 20, 6, 5, 28),
            (56, 136, 4, 0, 20, 6, 5, 28),
            (56, 136, 4, 3, 20, 0, 5, 28),
            (56, 136, 4, 3, 20, 6, 0, 28),
            (56, 136, usize::MAX, 3, 20, 6, 5, 28),
            (56, 136, 4, usize::MAX, 20, 6, 5, 28),
        ] {
            let input = vec![37; input_len];
            let mut output = vec![0xa5; output_len];
            assert_eq!(scale_rgba_bytes(&input, &mut output, width, rows, pitch, ow, oh, op), None);
            assert!(output.iter().all(|&byte| byte == 0xa5));
        }
    }

    fn object(fields: &[(&str, Value)]) -> Value {
        let mut object = Object::new();
        object.properties.extend(fields.iter().map(|(name, value)| ((*name).into(), value.clone())));
        Value::Object(heap::alloc(object))
    }

    fn surface(width: i32, height: i32, pitch: i32, pointer: i32) -> Value {
        let format = object(&[("BitsPerPixel", Value::I32(32)),
            ("Rmask", Value::I64(255)), ("Gmask", Value::I64(65280)),
            ("Bmask", Value::I64(16711680)), ("Amask", Value::I64(4278190080))]);
        object(&[("format", format), ("w", Value::I32(width)), ("h", Value::I32(height)),
            ("pitch", Value::I32(pitch)), ("pixels", Value::I32(pointer))])
    }

    fn invoke(seed: Vec<u8>, operation: impl Fn(&mut HostContext<'_>) -> Option<()> + Send + Sync + 'static)
        -> (bool, Vec<u8>)
    {
        let mut vm = VM::new();
        vm.memory.with_buffer_mut(|memory| *memory = seed);
        vm.register_host_fn("test", "scaled", Box::new(move |ctx, _|
            Value::I32(i32::from(operation(ctx).is_some()))));
        let mut chunk = Chunk::new("scaled-rgba-test");
        let index = chunk.add_import("test", "scaled");
        chunk.emit_call(index, 0, 0); chunk.emit_op(Op::RETURN, 0);
        let success = vm.run(vec![chunk]).unwrap() == Value::I32(1);
        (success, vm.memory.with_buffer(|memory| memory.to_vec()))
    }

    #[test]
    fn linear_disjoint_orders_and_unaligned_offsets_match_snapshot() {
        for (source, destination) in [(8usize, 200usize), (300, 0)] {
            let seed: Vec<u8> = (0usize..512).map(|i| (i.wrapping_mul(37) + 11) as u8).collect();
            let mut expected = seed.clone();
            for y in 0..5 { for x in 0..11 {
                let src = source + 3 + (y * 3 / 5) * 31 + (x * 7 / 11) * 4;
                let dst = destination + 7 + y * 49 + x * 4;
                expected[dst..dst + 4].copy_from_slice(&seed[src..src + 4]);
            }}
            let (success, actual) = invoke(seed, move |ctx|
                scale_linear_rgba_rows(ctx, &Value::I32(source as i32), &Value::I32(destination as i32),
                    3, 7, 7, 3, 31, 11, 5, 49));
            assert!(success); assert_eq!(actual, expected);
        }
    }

    #[test]
    fn overlapping_linear_ranges_reject_without_writes() {
        let seed: Vec<u8> = (0usize..512).map(|i| (i.wrapping_mul(37) + 11) as u8).collect();
        let (success, actual) = invoke(seed.clone(), |ctx|
            scale_linear_rgba_rows(ctx, &Value::I32(8), &Value::I32(24), 0, 0,
                7, 3, 31, 11, 5, 49));
        assert!(!success); assert_eq!(actual, seed);
    }

    #[test]
    fn registered_scaled_blit_preserves_existing_five_argument_abi() {
        let seed: Vec<u8> = (0usize..300).map(|i| (i.wrapping_mul(37) + 11) as u8).collect();
        let mut expected = seed.clone();
        for y in 0..5 { for x in 0..7 {
            let src = (y * 3 / 5) * 20 + (x * 4 / 7) * 4;
            let dst = 80 + y * 32 + x * 4;
            expected[dst..dst + 4].copy_from_slice(&seed[src..src + 4]);
        }}
        let mut vm = VM::new();
        vm.memory.with_buffer_mut(|memory| *memory = seed);
        register(&mut vm);
        let source = surface(4, 3, 20, 0);
        let destination = surface(7, 5, 32, 80);
        vm.register_host_fn("test", "source", Box::new(move |_, _| source.clone()));
        vm.register_host_fn("test", "destination", Box::new(move |_, _| destination.clone()));
        let mut chunk = Chunk::new("registered-scaled-rgba");
        let source = chunk.add_import("test", "source");
        let destination = chunk.add_import("test", "destination");
        let blit = chunk.add_import("libc:sdl", "convertBlit");
        chunk.emit_call(source, 0, 0); chunk.emit_ref_null(0x6f, 0);
        chunk.emit_call(destination, 0, 0); chunk.emit_ref_null(0x6f, 0);
        chunk.emit_i32_const(1, 0); chunk.emit_call(blit, 5, 0);
        chunk.emit_op(Op::RETURN, 0);
        assert_eq!(vm.run(vec![chunk]).unwrap(), Value::I32(0));
        assert_eq!(vm.memory.with_buffer(|memory| memory.to_vec()), expected);
    }

    #[test]
    fn cropped_adapter_rectangles_preserve_alpha_and_overlap_fallback() {
        for destination in [0usize, 300] {
            let seed: Vec<u8> = (0usize..800).map(|i| (i.wrapping_mul(37) + 11) as u8).collect();
            let mut expected = seed.clone();
            for y in 0..5 { for x in 0..7 {
                let src = (1 + y * 3 / 5) * 36 + (1 + x * 5 / 7) * 4;
                let dst = destination + (1 + y) * 48 + (2 + x) * 4;
                expected[dst..dst + 4].copy_from_slice(&seed[src..src + 4]);
            }}
            let rectangle = |x, y, w, h| object(&[("x", Value::I32(x)), ("y", Value::I32(y)),
                ("w", Value::I32(w)), ("h", Value::I32(h))]);
            let args = [surface(8, 6, 36, 0), rectangle(1, 1, 5, 3),
                surface(10, 9, 48, destination as i32), rectangle(2, 1, 7, 5), Value::I32(1)];
            let (success, actual) = invoke(seed, move |ctx| convert_blit(ctx, &args));
            assert!(success); assert_eq!(actual, expected);
        }
    }
}
