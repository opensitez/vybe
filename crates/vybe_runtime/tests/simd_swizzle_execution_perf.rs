//! Swizzle dispatch correctness and module-run timing, not kernel-only timing.
use vybe_runtime::{Chunk, Op, VM, Value};

fn vector(chunk: &mut Chunk, bytes: [u8; 16]) {
    chunk.emit_op(Op::V128_CONST, 0);
    for byte in bytes { chunk.emit(byte, 0); }
}

fn program(source: [u8; 16], indices: [u8; 16], count: usize) -> Chunk {
    let mut chunk = Chunk::new("swizzle-execution");
    vector(&mut chunk, source);
    for _ in 0..count {
        vector(&mut chunk, indices);
        chunk.emit_op(Op::I8X16_SWIZZLE, 0);
    }
    chunk.emit_op(Op::RETURN, 0);
    chunk
}

fn reference(mut source: [u8; 16], indices: [u8; 16], count: usize) -> [u8; 16] {
    for _ in 0..count {
        source = std::array::from_fn(|lane| {
            source.get(indices[lane] as usize).copied().unwrap_or(0)
        });
    }
    source
}

#[test]
fn swizzle_dispatch_preserves_indices_and_zeroing() {
    let source = std::array::from_fn(|lane| lane as u8 * 13 + 7);
    let mut selectors = vec![
        [15, 14, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1, 0],
        [0, 15, 16, 127, 128, 255, 1, 14, 17, 3, 0, 240, 13, 12, 11, 10],
    ];
    for index in 0..=255u16 { selectors.push([index as u8; 16]); }
    for indices in selectors {
        for count in [0, 1, 2, 17] {
            let actual = VM::new().run(vec![program(source, indices, count)]).unwrap();
            match actual {
                Value::V128(actual) => assert_eq!(actual, reference(source, indices, count)),
                other => panic!("expected vector, got {other:?}"),
            }
        }
    }
}

#[test]
#[ignore = "prepared bytecode swizzle module-run performance measurement"]
fn swizzle_module_run_benchmark() {
    use std::time::Instant;
    let source = std::array::from_fn(|lane| lane as u8 * 13 + 7);
    let indices = [15, 14, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1, 0];
    const COUNT: usize = 100_000;
    let expected = reference(source, indices, COUNT);
    let mut times = [0.0f64; 5];
    for sample in &mut times {
        let modules = vec![program(source, indices, COUNT)];
        let mut vm = VM::new();
        let start = Instant::now();
        let actual = vm.run(modules).unwrap();
        *sample = start.elapsed().as_secs_f64();
        match actual {
            Value::V128(actual) => assert_eq!(actual, expected),
            other => panic!("expected vector, got {other:?}"),
        }
    }
    times.sort_by(f64::total_cmp);
    eprintln!("swizzle module run: median {:.3} ms, {:.1} ns/iteration; samples {times:?}; includes VM::run setup/decode",
        times[2] * 1000.0, times[2] * 1e9 / COUNT as f64);
}
