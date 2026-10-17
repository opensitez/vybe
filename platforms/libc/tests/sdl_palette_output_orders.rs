include!("../src/emitter/sdl_runtime.rs");

#[test]
fn full_palette_output_orders_match_snapshot_for_doom_argb_and_scaling() {
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
    let source_format = object(&[("BitsPerPixel", Value::I32(8)), ("palette", palette),
        ("Rmask", Value::I32(0)), ("Gmask", Value::I32(0)),
        ("Bmask", Value::I32(0)), ("Amask", Value::I32(0))]);
    for masks in [[0xffu32, 0xff00, 0xff0000, 0xff000000],
        [0xff0000, 0xff00, 0xff, 0xff000000],
        [0xff000000, 0xff0000, 0xff00, 0xff]] {
        let destination_format = object(&[("BitsPerPixel", Value::I32(32)),
            ("Rmask", Value::I64(masks[0] as i64)), ("Gmask", Value::I64(masks[1] as i64)),
            ("Bmask", Value::I64(masks[2] as i64)), ("Amask", Value::I64(masks[3] as i64))]);
        for (ow, oh) in [(7usize, 3usize), (11, 5), (3, 2)] {
            for (source, destination) in [(8usize, 256usize), (700, 8), (8, 9)] {
                let sp = 11usize; let dp = ow * 4 + 5;
                let seed: Vec<u8> = (0usize..2048).map(|i| (i * 37 + 129) as u8).collect();
                let mut expected = seed.clone(); let mut actual = seed.clone();
                for y in 0..oh { for x in 0..ow {
                    let color = colors[usize::from(seed[source + (y * 3 / oh) * sp + x * 7 / ow])];
                    let mut packed = 0u32;
                    for component in 0..4 { packed |= u32::from(color[component]) << masks[component].trailing_zeros(); }
                    let start = destination + y * dp + x * 4;
                    expected[start..start + 4].copy_from_slice(&packed.to_le_bytes());
                }}
                let surface = |format, w, h, pitch, pointer| object(&[("format", format),
                    ("w", Value::I32(w)), ("h", Value::I32(h)), ("pitch", Value::I32(pitch)),
                    ("pixels", Value::I32(pointer))]);
                let args = [surface(source_format.clone(), 7, 3, sp as i32, source as i32), Value::Null,
                    surface(destination_format.clone(), ow as i32, oh as i32, dp as i32, destination as i32),
                    Value::Null, Value::I32(i32::from(ow != 7 || oh != 3))];
                let mut ctx = HostContext::empty(); ctx.memory = Some(&mut actual);
                assert_eq!(convert_blit(&mut ctx, &args), Some(()));
                assert_eq!(actual, expected, "masks={masks:?} {ow}x{oh} src={source} dst={destination}");
            }
        }
    }
}
