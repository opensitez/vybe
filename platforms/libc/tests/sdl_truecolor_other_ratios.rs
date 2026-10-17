include!("sdl_truecolor_double_frame.rs");

#[test]
#[ignore = "paired non-2x scaler regression control; not whole SDL or Doom timing"]
fn paired_non_double_ratio_benchmark() {
    use std::hint::black_box;
    use std::time::Instant;
    let (width, rows) = (320usize, 200usize);
    let pitch = width * 4 + 13;
    let input: Vec<u8> = (0..(rows - 1) * pitch + width * 4)
        .map(|index| (index * 73 + 91) as u8).collect();
    for (output_width, output_rows) in [(160usize, 100usize), (320, 200), (320, 400),
        (960, 600), (1280, 800)] {
        let output_pitch = output_width * 4 + 13;
        let span = (output_rows - 1) * output_pitch + output_width * 4;
        let mut samples = [Vec::new(), Vec::new()];
        for round in 0..19 {
            let order = if round % 2 == 0 { [0, 1] } else { [1, 0] };
            for path in order {
                let mut output = vec![0xa5; span];
                let start = Instant::now();
                for _ in 0..20 {
                    if path == 0 {
                        original_rgba_scale(black_box(&input), black_box(&mut output), width, rows, pitch,
                            output_width, output_rows, output_pitch).unwrap();
                    } else {
                        scale_rgba_bytes(black_box(&input), black_box(&mut output), width, rows, pitch,
                            output_width, output_rows, output_pitch).unwrap();
                    }
                }
                let elapsed = start.elapsed().as_secs_f64();
                let mut expected = vec![0xa5; span];
                original_rgba_scale(&input, &mut expected, width, rows, pitch,
                    output_width, output_rows, output_pitch).unwrap();
                assert_eq!(output, expected);
                if round >= 2 { samples[path].push(elapsed); }
            }
        }
        for sample in &mut samples { sample.sort_by(f64::total_cmp); }
        eprintln!("truecolor-other 320x200 -> {output_width}x{output_rows} original_ms={:.3} current_ms={:.3} original_over_current={:.3}",
            samples[0][8] * 1000.0, samples[1][8] * 1000.0, samples[0][8] / samples[1][8]);
    }
}
