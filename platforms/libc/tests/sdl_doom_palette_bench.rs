include!("../src/emitter/sdl_runtime.rs");

#[test]
#[ignore = "optimized ARGB palette adapter benchmark; not Doom FPS"]
fn doom_argb_palette_adapter_against_snapshot_pixel_encoding() {
    use std::{hint::black_box, time::Instant};
    assert!(!cfg!(debug_assertions));
    fn object(fields: &[(&str, Value)]) -> Value {
        let mut value = Object::new();
        value.properties.extend(fields.iter().map(|(key, value)| ((*key).into(), value.clone())));
        Value::Object(heap::alloc(value))
    }
    let colors: [[u8; 4]; 256] = std::array::from_fn(|i|
        [i as u8, (i * 37) as u8, (255 - i) as u8, (i * 13) as u8]);
    let mut entries = Object::new();
    entries.kind = ObjectKind::Array(colors.iter().map(|c| object(&[
        ("r", Value::I32(c[0] as i32)), ("g", Value::I32(c[1] as i32)),
        ("b", Value::I32(c[2] as i32)), ("a", Value::I32(c[3] as i32)),
    ])).collect());
    let palette = object(&[("colors", Value::Object(heap::alloc(entries))), ("ncolors", Value::I32(256))]);
    let indexed = object(&[("BitsPerPixel", Value::I32(8)), ("palette", palette),
        ("Rmask", Value::I32(0)), ("Gmask", Value::I32(0)),
        ("Bmask", Value::I32(0)), ("Amask", Value::I32(0))]);
    let masks = [0xff0000u32, 0xff00, 0xff, 0xff000000];
    let argb = object(&[("BitsPerPixel", Value::I32(32)),
        ("Rmask", Value::I64(masks[0] as i64)), ("Gmask", Value::I64(masks[1] as i64)),
        ("Bmask", Value::I64(masks[2] as i64)), ("Amask", Value::I64(masks[3] as i64))]);
    let pixel_format = PixelFormat { bytes: 4, masks, palette: None };
    let (width, rows, source, sp) = (320usize, 200usize, 8usize, 323usize);
    let destination = rows * sp + 64;
    let surface = |format, w, h, pitch, address| object(&[("format", format),
        ("w", Value::I32(w)), ("h", Value::I32(h)), ("pitch", Value::I32(pitch)),
        ("pixels", Value::I32(address))]);
    for (ow, oh) in [(320usize, 200usize), (640, 400), (960, 600)] {
        let dp = ow * 4 + 13;
        let args = [surface(indexed.clone(), width as i32, rows as i32, sp as i32, source as i32),
            Value::Null, surface(argb.clone(), ow as i32, oh as i32, dp as i32, destination as i32),
            Value::Null, Value::I32(i32::from(ow != width || oh != rows))];
        let seed: Vec<u8> = (0usize..destination + oh * dp + 11)
            .map(|i| (i * 37 + 129) as u8).collect();
        let mut reference = seed.clone(); let mut candidate = seed;
        let measure = |memory: &mut Vec<u8>, fast: bool, iterations: usize| {
            let mut ctx = HostContext::empty(); ctx.memory = Some(memory);
            let start = Instant::now();
            for _ in 0..iterations {
                if black_box(fast) {
                    convert_blit(&mut ctx, black_box(&args)).unwrap();
                } else {
                    // Omits adapter lookup/palette-decoding overhead, favoring
                    // the old operations. Snapshot ranges match conversion.
                    ctx.with_linear_memory_mut(|memory| {
                        let length = (rows - 1) * sp + width;
                        let input = memory[source..source + length].to_vec();
                        let length = (oh - 1) * dp + ow * 4;
                        let mut output = memory[destination..destination + length].to_vec();
                        for y in 0..oh { for x in 0..ow {
                            let index = input[(y * rows / oh) * sp + x * width / ow];
                            let start = y * dp + x * 4;
                            black_box(&pixel_format).encode(colors[usize::from(index)],
                                &mut output[start..start + 4]).unwrap();
                        }}
                        memory[destination..destination + length].copy_from_slice(&output);
                    }).unwrap();
                }
            }
            start.elapsed().as_nanos() as f64
        };
        measure(&mut reference, false, 1); measure(&mut candidate, true, 1);
        assert_eq!(candidate, reference);
        let mut ratios = Vec::new();
        for round in 0..19 {
            let (old, new) = if round % 2 == 0 {
                (measure(&mut reference, false, 10), measure(&mut candidate, true, 10))
            } else {
                let new = measure(&mut candidate, true, 10);
                (measure(&mut reference, false, 10), new)
            };
            if round >= 2 { ratios.push(old / new); }
        }
        assert_eq!(candidate, reference);
        ratios.sort_by(f64::total_cmp);
        println!("ARGB palette adapter 320x200 -> {ow}x{oh}: {:.3}x", ratios[ratios.len()/2]);
    }
}
