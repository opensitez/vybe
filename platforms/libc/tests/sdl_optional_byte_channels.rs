include!("../src/emitter/sdl_runtime.rs");

#[test]
fn omitted_channels_are_rejected_or_match_existing_defaults() {
    for missing in 0..4 {
        let mut masks = [0u32; 4]; let mut byte = 0;
        for component in 0..4 {
            if component != missing {
                masks[component] = 0xff << (byte * 8); byte += 1;
            }
        }
        let sf = PixelFormat { bytes: 3, masks, palette: None };
        for destination_bytes in [3usize, 4] {
            let df = PixelFormat { bytes: destination_bytes,
                masks: [0xff, 0xff00, 0xff0000, if destination_bytes == 4 { 0xff000000 } else { 0 }],
                palette: None };
            let mut memory = vec![0xa5; 128]; memory[8..11].copy_from_slice(&[17, 53, 211]);
            let before = memory.clone(); let mut expected = memory.clone();
            let reader = HostContext::empty();
            let color = sf.decode(&reader, &before, 8).unwrap();
            df.encode(color, &mut expected[64..64 + destination_bytes]).unwrap();
            let accepted = sf.byte_channels();
            let mut ctx = HostContext::empty(); ctx.memory = Some(&mut memory);
            let result = convert_linear_byte_rows(&mut ctx, &Value::I32(8), &Value::I32(64),
                0, 0, 1, 1, 3, destination_bytes, &sf, &df);
            println!("missing component={missing} destination_bytes={destination_bytes} fast_proof={accepted}");
            if accepted { assert_eq!(result, Some(())); assert_eq!(memory, expected); }
            else { assert_eq!(result, None); assert_eq!(memory, before); }
        }
    }
}
