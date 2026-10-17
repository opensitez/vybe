include!("../src/emitter/sdl_runtime.rs");

#[cfg(test)]
mod blit_host_tests {
    use super::*;
    use vybe_runtime::{Chunk, opcode::Op};

    fn object(fields: &[(&str, Value)]) -> Value {
        let mut value = Object::new();
        value.properties.extend(fields.iter().map(|(name, value)| ((*name).into(), value.clone())));
        Value::Object(heap::alloc(value))
    }

    fn rgba_format() -> Value {
        object(&[("BitsPerPixel", Value::I32(32)),
            ("Rmask", Value::I64(255)), ("Gmask", Value::I64(65280)),
            ("Bmask", Value::I64(16711680)), ("Amask", Value::I64(4278190080))])
    }

    fn surface(format: Value, width: i32, height: i32, pitch: i32, pointer: i32) -> Value {
        object(&[("format", format), ("w", Value::I32(width)), ("h", Value::I32(height)),
            ("pitch", Value::I32(pitch)), ("pixels", Value::I32(pointer))])
    }

    fn invoke(memory: Vec<u8>, source: Value, destination: Value) -> Vec<u8> {
        let mut vm = VM::new();
        vm.memory.with_buffer_mut(|bytes| *bytes = memory);
        register(&mut vm);
        vm.register_host_fn("test", "source", Box::new(move |_, _| source.clone()));
        vm.register_host_fn("test", "destination", Box::new(move |_, _| destination.clone()));
        let mut chunk = Chunk::new("registered-sdl-blit");
        let source = chunk.add_import("test", "source");
        let destination = chunk.add_import("test", "destination");
        let blit = chunk.add_import("libc:sdl", "convertBlit");
        chunk.emit_call(source, 0, 0);
        chunk.emit_ref_null(0x6f, 0);
        chunk.emit_call(destination, 0, 0);
        chunk.emit_ref_null(0x6f, 0);
        chunk.emit_i32_const(0, 0);
        chunk.emit_call(blit, 5, 0);
        chunk.emit_op(Op::RETURN, 0);
        assert_eq!(vm.run(vec![chunk]).unwrap(), Value::I32(0));
        vm.memory.with_buffer(|bytes| bytes.to_vec())
    }

    #[test]
    fn registered_rgba_blit_keeps_overlap_and_padding() {
        let memory: Vec<u8> = (0..96).map(|i| (i * 37 + 11) as u8).collect();
        let mut expected = memory.clone();
        for row in 0..2 {
            expected[8 + row * 24..24 + row * 24]
                .copy_from_slice(&memory[row * 24..16 + row * 24]);
        }
        let actual = invoke(memory, surface(rgba_format(), 4, 2, 24, 0),
            surface(rgba_format(), 4, 2, 24, 8));
        assert_eq!(actual, expected);
    }

    #[test]
    fn registered_palette_blit_keeps_unsigned_indices_and_exact_output() {
        let mut colors = Object::new();
        colors.kind = ObjectKind::Array((0..256).map(|i|
            Value::I32((i << 16) | ((255 - i) << 8) | ((i * 7) & 255))).collect());
        let palette = object(&[("colors", Value::Object(heap::alloc(colors))),
            ("ncolors", Value::I32(256))]);
        let indexed = object(&[("BitsPerPixel", Value::I32(8)), ("palette", palette),
            ("Rmask", Value::I32(0)), ("Gmask", Value::I32(0)),
            ("Bmask", Value::I32(0)), ("Amask", Value::I32(0))]);
        let mut memory = vec![0xa5; 64];
        let indices = [0u8, 127, 128, 255];
        memory[..4].copy_from_slice(&indices);
        let mut expected = memory.clone();
        for (x, index) in indices.into_iter().enumerate() {
            expected[16 + x * 4..20 + x * 4]
                .copy_from_slice(&[index, 255 - index, index.wrapping_mul(7), 255]);
        }
        let actual = invoke(memory, surface(indexed, 4, 1, 4, 0),
            surface(rgba_format(), 4, 1, 16, 16));
        assert_eq!(actual, expected);
    }
}
