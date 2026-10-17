include!("../src/emitter/sdl_runtime.rs");

fn original_packed_format(ctx: &HostContext<'_>, value: &Value) -> Option<PixelFormat> {
    if matches!(value, Value::Object(_)) { return PixelFormat::read(ctx, value); }
    let code = value.as_i64() as u32;
    let kind = (code >> 24) & 15;
    if !(4..=6).contains(&kind) { return PixelFormat::read(ctx, value); }
    let bytes = (code & 255) as usize;
    if !(1..=4).contains(&bytes) { return None; }
    let widths: &[u32] = match (code >> 16) & 15 {
        1 => &[3, 3, 2], 2 => &[4, 4, 4, 4], 3 => &[1, 5, 5, 5],
        4 => &[5, 5, 5, 1], 5 => &[5, 6, 5], 6 => &[8, 8, 8, 8],
        7 => &[2, 10, 10, 10], 8 => &[10, 10, 10, 2], _ => return None,
    };
    let channels = match (code >> 20) & 15 {
        1 => [4, 0, 1, 2], 2 => [0, 1, 2, 4], 3 => [3, 0, 1, 2],
        4 => [0, 1, 2, 3], 5 => [4, 2, 1, 0], 6 => [2, 1, 0, 4],
        7 => [3, 2, 1, 0], 8 => [2, 1, 0, 3], _ => return None,
    };
    let channels: Vec<_> = channels.into_iter()
        .filter(|&channel| widths.len() == 4 || channel != 4).collect();
    if channels.len() != widths.len() { return None; }
    let mut masks = [0u32; 4];
    let mut shift = widths.iter().sum::<u32>();
    for (&channel, &width) in channels.iter().zip(widths) {
        shift -= width;
        if channel < 4 { masks[channel] = ((1u32 << width) - 1) << shift; }
    }
    Some(PixelFormat { bytes, masks, palette: None })
}

#[test]
fn stack_format_decoder_matches_original_packed_numeric_codes() {
    let ctx = HostContext::empty();
    for marker in [0x1000_0000u32, 0xf000_0000] {
        for kind in 4..=6 {
            for order in 0..16 {
                for layout in 0..16 {
                    for bytes in 0..=5 {
                        let code = marker | kind << 24 | order << 20 | layout << 16 | 32 << 8 | bytes;
                        for value in [Value::I32(code as i32), Value::I64(i64::from(code)), Value::F64(f64::from(code))] {
                            let expected = original_packed_format(&ctx, &value)
                                .map(|format| (format.bytes, format.masks));
                            let actual = PixelFormat::read(&ctx, &value)
                                .map(|format| (format.bytes, format.masks));
                            assert_eq!(actual, expected, "code={code:#x} value={value:?}");
                        }
                    }
                }
            }
        }
    }
}

#[test]
#[ignore = "paired numeric pixel-format allocation benchmark; not frame or Doom timing"]
fn paired_numeric_format_decoding_benchmark() {
    use std::hint::black_box;
    use std::time::Instant;
    let ctx = HostContext::empty();
    let codes = [Value::I64(0x1636_2004), Value::I64(0x1646_2004),
        Value::I64(0x1515_1002), Value::I64(0x1616_2004)];
    let mut samples = [Vec::new(), Vec::new()];
    for round in 0..19 {
        let order = if round % 2 == 0 { [0, 1] } else { [1, 0] };
        for path in order {
            let start = Instant::now();
            let mut checksum = 0u64;
            for index in 0..50_000 {
                let value = black_box(&codes[index % codes.len()]);
                let format = if path == 0 {
                    original_packed_format(&ctx, value)
                } else {
                    PixelFormat::read(&ctx, value)
                }.unwrap();
                checksum += u64::from(format.masks[0]);
                black_box(format);
            }
            black_box(checksum);
            let elapsed = start.elapsed().as_secs_f64();
            if round >= 2 { samples[path].push(elapsed); }
        }
    }
    for sample in &mut samples { sample.sort_by(f64::total_cmp); }
    eprintln!("numeric-format calls=50000 original_ms={:.3} stack_ms={:.3} original_over_stack={:.3}",
        samples[0][8] * 1000.0, samples[1][8] * 1000.0, samples[0][8] / samples[1][8]);
}
