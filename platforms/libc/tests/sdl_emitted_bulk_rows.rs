use vybe_platform_libc::emitter::{sdl, sdl_runtime};
use vybe_runtime::{Chunk, VM, Value, heap, opcode::Op, value::Object};

fn object(fields: &[(&str, Value)]) -> Value {
    let mut object = Object::new();
    object.properties.extend(fields.iter().map(|(key, value)| ((*key).into(), value.clone())));
    Value::Object(heap::alloc(object))
}

fn surface(bytes: usize, width: usize, rows: usize, pitch: usize, pointer: usize) -> Value {
    let format = object(&[("BitsPerPixel", Value::F64((bytes * 8) as f64)),
        ("Rmask", Value::F64(255.0)), ("Gmask", Value::F64(65280.0)),
        ("Bmask", Value::F64(16711680.0)),
        ("Amask", Value::F64(if bytes == 4 { 4278190080.0 } else { 0.0 }))]);
    object(&[("format", format), ("w", Value::F64(width as f64)),
        ("h", Value::F64(rows as f64)), ("pitch", Value::F64(pitch as f64)),
        ("pixels", Value::F64(pointer as f64))])
}

#[test]
fn emitted_bulk_and_pitched_copies_preserve_snapshot_and_padding() {
    for bytes in [3usize, 4] { for padding in [0usize, 5] {
      for rows in [1usize, 3, 5] {
        let (width, source_address) = (4usize, 64usize);
        let pitch = width * bytes + padding;
        for displacement in [-1isize, 0, 1, pitch as isize + 1, 128] {
            let destination_address = source_address.checked_add_signed(displacement).unwrap();
            let seed: Vec<u8> = (0usize..512).map(|i| (i * 37 + 129) as u8).collect();
            let mut expected = seed.clone();
            for row in 0..rows {
                let source = source_address + row * pitch;
                let destination = destination_address + row * pitch;
                expected[destination..destination + width * bytes]
                    .copy_from_slice(&seed[source..source + width * bytes]);
            }
            let source = surface(bytes, width, rows, pitch, source_address);
            let destination = surface(bytes, width, rows, pitch, destination_address);
            let mut vm = VM::new(); vm.memory.with_buffer_mut(|memory| *memory = seed);
            vybe_runtime::js_builtins::register(&mut vm);
            sdl_runtime::register(&mut vm);
            vm.register_host_fn("test", "source", Box::new(move |_, _| source.clone()));
            vm.register_host_fn("test", "destination", Box::new(move |_, _| destination.clone()));
            let mut chunks = vec![Chunk::new("emitted-sdl-bulk-rows")];
            let src = chunks[0].add_import("test", "source");
            let dst = chunks[0].add_import("test", "destination");
            chunks[0].emit_call(src, 0, 0); chunks[0].emit_ref_null(0x6f, 0);
            chunks[0].emit_call(dst, 0, 0); chunks[0].emit_ref_null(0x6f, 0);
            sdl::emit_sdl_lower_blit(&mut chunks, 0, 4, 0);
            chunks[0].emit_op(Op::RETURN, 0);
            assert_eq!(vm.run(chunks).unwrap(), Value::I32(0));
            assert_eq!(vm.memory.with_buffer(|memory| memory.to_vec()), expected,
                "bytes={bytes} padding={padding} displacement={displacement}");
        }
    }}}
}
