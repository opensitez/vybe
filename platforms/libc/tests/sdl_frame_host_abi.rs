// Compile and exercise the actual SDL host adapter through the VM call ABI.
include!("../src/emitter/sdl_runtime.rs");

#[cfg(test)]
mod registered_frame_tests {
    use super::*;
    use vybe_runtime::{Chunk, opcode::Op};

    fn palette_with_four_colors() -> Value {
        let mut colors = Object::new();
        colors.kind = ObjectKind::Array(vec![Value::I32(0); 4]);
        let mut palette = Object::new();
        palette.properties.insert("ncolors".into(), Value::I32(4));
        palette.properties.insert("colors".into(), Value::Object(heap::alloc(colors)));
        palette.properties.insert("__sdl_first".into(), Value::I32(1));
        Value::Object(heap::alloc(palette))
    }

    fn color_bytes(ctx: &HostContext<'_>, palette: &Value, index: usize) -> [u8; 4] {
        let colors = palette_source(Some(ctx), Some(palette)).unwrap();
        let color = indexed_value(Some(ctx), Some(&colors), index).unwrap();
        ["r", "g", "b", "a"].map(|key| field(ctx, &color, key).unwrap().as_i32() as u8)
    }

    #[test]
    fn linear_color_updates_copy_rgba_and_do_not_alias_the_callers_buffer() {
        for address in [Value::I32(4), Value::I64(4), Value::F64(4.0)] {
            let mut memory = vec![9, 9, 9, 9, 10, 20, 30, 40, 50, 60, 70, 80];
            let mut ctx = HostContext::empty();
            ctx.memory = Some(&mut memory);
            let palette = palette_with_four_colors();
            assert_eq!(set_palette_colors(&ctx,
                &[palette.clone(), address, Value::I32(1), Value::I32(2)]), Some(()));
            ctx.with_linear_memory_mut(|bytes| bytes.fill(0));
            assert_eq!(color_bytes(&ctx, &palette, 1), [10, 20, 30, 40]);
            assert_eq!(color_bytes(&ctx, &palette, 2), [50, 60, 70, 80]);
        }
    }

    #[test]
    fn wider_typed_color_storage_is_copied_as_bytes_at_its_view_offset() {
        let ctx = HostContext::empty();
        let source = typed_byte_array(vec![9, 9, 9, 9, 10, 20, 30, 40, 50, 60, 70, 80],
            TypedElemKind::U8, "Uint8Array");
        let Value::Object(object) = &source else { unreachable!() };
        let mut object = object.lock().unwrap();
        let ObjectKind::TypedArray(array) = &mut object.kind else { unreachable!() };
        array.elem = TypedElemKind::I32;
        array.byte_offset = 4;
        array.length = 2;
        drop(object);
        let palette = palette_with_four_colors();
        assert_eq!(set_palette_colors(&ctx,
            &[palette.clone(), source.clone(), Value::I32(1), Value::I32(2)]), Some(()));
        let Value::Object(object) = source else { unreachable!() };
        let object = object.lock().unwrap();
        let ObjectKind::TypedArray(array) = &object.kind else { unreachable!() };
        array.buffer.lock().unwrap().fill(0);
        drop(object);
        assert_eq!(color_bytes(&ctx, &palette, 1), [10, 20, 30, 40]);
        assert_eq!(color_bytes(&ctx, &palette, 2), [50, 60, 70, 80]);
    }

    #[test]
    fn empty_and_invalid_color_updates_do_not_mutate_the_palette() {
        let mut memory = vec![0; 16];
        let mut ctx = HostContext::empty();
        ctx.memory = Some(&mut memory);
        let palette = palette_with_four_colors();
        let Value::Object(object) = &palette else { unreachable!() };
        let before = object.lock().unwrap().properties.clone();
        assert_eq!(set_palette_colors(&ctx,
            &[palette.clone(), Value::Null, Value::I32(2), Value::I32(0)]), Some(()));
        for (source, first, count) in [
            (Value::I32(4), 3, 2), (Value::I32(15), 0, 1),
            (Value::I32(-4), 0, 1), (Value::F64(f64::NAN), 0, 1),
            (Value::F64(4.5), 0, 1), (Value::I32(0), 0, 1),
            (Value::I32(4), 5, 0),
        ] {
            assert_eq!(set_palette_colors(&ctx,
                &[palette.clone(), source, Value::I32(first), Value::I32(count)]), None);
        }
        assert_eq!(object.lock().unwrap().properties, before);
    }

    fn image_bytes(image: &Value) -> Vec<u8> {
        let ctx = HostContext::empty();
        let data = field(&ctx, image, "data").unwrap();
        bytes_from_value(Some(&ctx), Some(&data)).unwrap()
    }

    #[test]
    fn rgba_host_call_preserves_alpha_padding_and_channel_order() {
        for (order, pitch, expected) in [
            (2, 8, vec![200, 100, 50, 7, 0, 255, 127, 128,
                11, 22, 33, 44, 55, 66, 77, 88]),
            (2, 12, vec![200, 100, 50, 7, 0, 255, 127, 128,
                11, 22, 33, 44, 55, 66, 77, 88]),
            (5, 12, vec![50, 100, 200, 7, 127, 255, 0, 128,
                33, 22, 11, 44, 77, 66, 55, 88]),
        ] {
            // End exactly at the last pixel, without unused final-row padding.
            let mut memory = vec![0; 1 + pitch as usize + 8];
            memory[1..9].copy_from_slice(&[200, 100, 50, 7, 0, 255, 127, 128]);
            let second_row = 1 + pitch as usize;
            memory[second_row..second_row + 8].copy_from_slice(&[11, 22, 33, 44, 55, 66, 77, 88]);
            let mut vm = VM::new();
            vm.memory.with_buffer_mut(|bytes| *bytes = memory);
            register(&mut vm);
            let mut chunk = Chunk::new("sdl_rgba_host_abi");
            let call = chunk.add_import("libc:sdl", "rgbaImageData");
            let format = (7 << 24) | (order << 20) | (32 << 8) | 4;
            for argument in [1, 2, 2, format, pitch] { chunk.emit_i32_const(argument, 0); }
            chunk.emit_call(call, 5, 0);
            chunk.emit_op(Op::RETURN, 0);
            let image = vm.run(vec![chunk]).unwrap();
            assert_eq!(image_bytes(&image), expected, "order={order} pitch={pitch}");
        }
    }

    #[test]
    fn linear_palette_host_call_uses_actual_bytes_not_black_defaults() {
        let mut palette = Object::new();
        palette.kind = ObjectKind::Array(vec![Value::I32(0x123456), Value::I32(0xabcdef)]);
        let palette = Value::Object(heap::alloc(palette));
        let mut vm = VM::new();
        vm.memory.with_buffer_mut(|bytes| *bytes = vec![0, 1, 2, 255]);
        register(&mut vm);
        // Test-only producer supplies the same managed palette shape as SDL.
        vm.register_host_fn("test", "palette", Box::new(move |_, _| palette.clone()));
        let mut chunk = Chunk::new("sdl_palette_host_abi");
        let palette = chunk.add_import("test", "palette");
        let present = chunk.add_import("libc:sdl", "palettedImageData");
        chunk.emit_i32_const(0, 0);
        chunk.emit_call(palette, 0, 0);
        chunk.emit_i32_const(4, 0);
        chunk.emit_i32_const(1, 0);
        chunk.emit_call(present, 4, 0);
        chunk.emit_op(Op::RETURN, 0);
        let image = vm.run(vec![chunk]).unwrap();
        assert_eq!(image_bytes(&image), [0x12, 0x34, 0x56, 255, 0xab, 0xcd, 0xef, 255,
            0, 0, 0, 255, 0, 0, 0, 255]);
    }

    #[test]
    fn malformed_pitch_is_rejected_through_the_registered_host_call() {
        for width in [2, i32::MAX] {
            let mut vm = VM::new();
            vm.memory.with_buffer_mut(|bytes| *bytes = vec![0; 16]);
            register(&mut vm);
            let mut chunk = Chunk::new("sdl_bad_pitch");
            let call = chunk.add_import("libc:sdl", "rgbaImageData");
            let format = (7 << 24) | (2 << 20) | (32 << 8) | 4;
            for argument in [0, width, 1, format, 0] { chunk.emit_i32_const(argument, 0); }
            chunk.emit_call(call, 5, 0);
            chunk.emit_op(Op::RETURN, 0);
            assert!(matches!(vm.run(vec![chunk]).unwrap(), Value::Null));
        }
    }
}
