include!("../src/emitter/sdl_runtime.rs");

#[cfg(test)]
mod row_copy_benchmark {
    use super::*;
    use std::{hint::black_box, time::Instant};

    // Conservative reference: two checked snapshots and row copying under
    // one borrow, without the old converter's per-pixel decode/encode work.
    fn snapshot(ctx: &mut HostContext<'_>, source: usize, destination: usize,
        bytes: usize, rows: usize, pitch: usize) -> Option<()> {
        let span = (rows - 1).checked_mul(pitch)?.checked_add(bytes)?;
        let source_end = source.checked_add(span)?;
        let destination_end = destination.checked_add(span)?;
        ctx.with_linear_memory_mut(|memory| {
            if source_end > memory.len() || destination_end > memory.len() { return None; }
            let input = memory[source..source_end].to_vec();
            let mut output = memory[destination..destination_end].to_vec();
            for row in 0..rows {
                let start = row * pitch;
                output[start..start + bytes].copy_from_slice(&input[start..start + bytes]);
            }
            memory[destination..destination_end].copy_from_slice(&output);
            Some(())
        })?
    }

    fn bulk(ctx: &mut HostContext<'_>, source: usize, destination: usize,
        bytes: usize, rows: usize, pitch: usize) -> Option<()> {
        copy_linear_rgba_rows(ctx, &Value::I32(source as i32),
            &Value::I32(destination as i32), 0, 0, bytes, rows, pitch)
    }

    fn unequal_snapshot(ctx: &mut HostContext<'_>, source: usize, destination: usize,
        bytes: usize, rows: usize, sp: usize, dp: usize) -> Option<()> {
        let source_span = (rows - 1).checked_mul(sp)?.checked_add(bytes)?;
        let destination_span = (rows - 1).checked_mul(dp)?.checked_add(bytes)?;
        let source_end = source.checked_add(source_span)?;
        let destination_end = destination.checked_add(destination_span)?;
        ctx.with_linear_memory_mut(|memory| {
            if source_end > memory.len() || destination_end > memory.len() { return None; }
            let input = memory[source..source_end].to_vec();
            let mut output = memory[destination..destination_end].to_vec();
            for row in 0..rows {
                output[row * dp..row * dp + bytes].copy_from_slice(&input[row * sp..row * sp + bytes]);
            }
            memory[destination..destination_end].copy_from_slice(&output);
            Some(())
        })?
    }

    fn original_row_loop(ctx: &mut HostContext<'_>, source: usize, destination: usize,
        bytes: usize, rows: usize, pitch: usize) -> Option<()> {
        let span = (rows - 1).checked_mul(pitch)?.checked_add(bytes)?;
        let source_end = source.checked_add(span)?;
        let destination_end = destination.checked_add(span)?;
        ctx.with_linear_memory_mut(|memory| {
            if source_end > memory.len() || destination_end > memory.len() { return None; }
            if source == destination { return Some(()); }
            for step in 0..rows {
                let row = if destination > source { rows - 1 - step } else { step };
                let offset = row * pitch;
                memory.copy_within(source + offset..source + offset + bytes, destination + offset);
            }
            Some(())
        })?
    }

    #[test]
    #[ignore = "optimized packed-span coalescing benchmark, not Doom FPS"]
    fn paired_packed_span_against_original_row_loop() {
        assert!(!cfg!(debug_assertions));
        type CopyRows = fn(&mut HostContext<'_>, usize, usize, usize, usize, usize) -> Option<()>;
        for (width, rows, pixel_bytes) in [(320usize, 200usize, 3usize), (320, 200, 4), (640, 480, 4)] {
            let bytes = width * pixel_bytes; let span = rows * bytes; let source = 64usize;
            for overlap in [false, true] {
                let destination = if overlap { source + 1 } else { source + span + 64 };
                let seed: Vec<u8> = (0usize..destination + span + 11)
                    .map(|i| (i.wrapping_mul(37) + 129) as u8).collect();
                let mut reference = seed.clone(); let mut candidate = seed;
                let measure = |memory: &mut Vec<u8>, operation: CopyRows, iterations: usize| {
                    let mut ctx = HostContext::empty(); ctx.memory = Some(memory);
                    let start = Instant::now();
                    for _ in 0..iterations {
                        black_box(operation)(&mut ctx, source, destination, bytes, rows, bytes).unwrap();
                    }
                    start.elapsed().as_nanos() as f64
                };
                measure(&mut reference, original_row_loop, 1);
                measure(&mut candidate, bulk, 1);
                assert_eq!(candidate, reference);
                let mut ratios = Vec::new();
                for round in 0..19 {
                    let (old, new) = if round % 2 == 0 {
                        (measure(&mut reference, original_row_loop, 50), measure(&mut candidate, bulk, 50))
                    } else {
                        let new = measure(&mut candidate, bulk, 50);
                        (measure(&mut reference, original_row_loop, 50), new)
                    };
                    if round >= 2 { ratios.push(old / new); }
                }
                assert_eq!(candidate, reference);
                ratios.sort_by(f64::total_cmp);
                println!("packed coalescing {width}x{rows} bytes={pixel_bytes} overlap={overlap}: {:.3}x",
                    ratios[ratios.len()/2]);
            }
        }
    }

    fn unequal_bulk(ctx: &mut HostContext<'_>, source: usize, destination: usize,
        bytes: usize, rows: usize, sp: usize, dp: usize) -> Option<()> {
        copy_linear_matching_rows(ctx, &Value::I32(source as i32),
            &Value::I32(destination as i32), 0, 0, bytes, rows, sp, dp)
    }

    #[test]
    #[ignore = "optimized paired unequal-pitch experiment; not a frame benchmark"]
    fn paired_unequal_pitch_copy_in_both_address_orders() {
        assert!(!cfg!(debug_assertions));
        type CopyRows = fn(&mut HostContext<'_>, usize, usize, usize, usize, usize, usize) -> Option<()>;
        for (width, rows, pixel_bytes) in [(320usize, 200usize, 3usize), (320, 200, 4), (640, 480, 4)] {
            let bytes = width * pixel_bytes; let sp = bytes + 3; let dp = bytes + 17;
            let span = rows * sp.max(dp);
            for reverse in [false, true] {
                let (source, destination) = if reverse { (span + 64, 8) } else { (8, span + 64) };
                let seed: Vec<u8> = (0usize..2 * span + 128)
                    .map(|i| (i.wrapping_mul(37) + 129) as u8).collect();
                let mut reference = seed.clone(); let mut candidate = seed;
                let measure = |memory: &mut Vec<u8>, operation: CopyRows, iterations: usize| {
                    let mut ctx = HostContext::empty(); ctx.memory = Some(memory);
                    let start = Instant::now();
                    for _ in 0..iterations {
                        black_box(operation)(&mut ctx, source, destination, bytes, rows, sp, dp).unwrap();
                    }
                    start.elapsed().as_nanos() as f64
                };
                measure(&mut reference, unequal_snapshot, 1);
                measure(&mut candidate, unequal_bulk, 1);
                assert_eq!(candidate, reference);
                let mut ratios = Vec::new();
                for round in 0..19 {
                    let (old, new) = if round % 2 == 0 {
                        (measure(&mut reference, unequal_snapshot, 30), measure(&mut candidate, unequal_bulk, 30))
                    } else {
                        let new = measure(&mut candidate, unequal_bulk, 30);
                        (measure(&mut reference, unequal_snapshot, 30), new)
                    };
                    if round >= 2 { ratios.push(old / new); }
                }
                assert_eq!(candidate, reference);
                ratios.sort_by(f64::total_cmp);
                println!("unequal-pitch {width}x{rows} bytes={pixel_bytes} reverse={reverse}: {:.3}x",
                    ratios[ratios.len()/2]);
            }
        }
    }

    #[test]
    #[ignore = "optimized paired performance experiment; not a frame benchmark"]
    fn paired_snapshot_and_actual_sdl_row_copier() {
        assert!(!cfg!(debug_assertions), "compile with optimizations");
        type CopyRows = fn(&mut HostContext<'_>, usize, usize, usize, usize, usize) -> Option<()>;
        for (width, rows, pixel_bytes) in [(320usize, 200usize, 3usize), (320, 200, 4), (640, 480, 4)] {
            for padding in [0usize, 17] { for overlap in [false, true] {
                let bytes = width * pixel_bytes;
                let pitch = bytes + padding;
                let span = (rows - 1) * pitch + bytes;
                let source = 256usize;
                let destination = if overlap { source + 1 } else { source + span + 64 };
                let seed: Vec<u8> = (0..destination + span + 11)
                    .map(|i: usize| (i.wrapping_mul(37) + 129) as u8).collect();
                let mut expected = seed.clone();
                let mut actual = seed.clone();
                for (memory, operation) in [(&mut expected, snapshot as CopyRows), (&mut actual, bulk)] {
                    let mut ctx = HostContext::empty(); ctx.memory = Some(memory);
                    operation(&mut ctx, source, destination, bytes, rows, pitch).unwrap();
                }
                assert_eq!(actual, expected);
                let mut reference_memory = seed.clone();
                let mut candidate_memory = seed;
                let mut ratios = Vec::new();
                for round in 0..19 {
                    let measure = |memory: &mut Vec<u8>, operation: CopyRows| {
                        let mut ctx = HostContext::empty(); ctx.memory = Some(memory);
                        let start = Instant::now();
                        for _ in 0..30 {
                            black_box(operation)(&mut ctx, source, destination, bytes, rows, pitch).unwrap();
                        }
                        start.elapsed().as_nanos() as f64
                    };
                    let (old, new) = if round % 2 == 0 {
                        (measure(&mut reference_memory, snapshot), measure(&mut candidate_memory, bulk))
                    } else {
                        let new = measure(&mut candidate_memory, bulk);
                        (measure(&mut reference_memory, snapshot), new)
                    };
                    if round >= 2 { ratios.push(old / new); }
                }
                assert_eq!(candidate_memory, reference_memory);
                ratios.sort_by(f64::total_cmp);
                println!("SDL rows {width}x{rows} bytes={pixel_bytes} padding={padding} overlap={overlap}: {:.3}x",
                    ratios[ratios.len() / 2]);
            }}
        }
    }
}
