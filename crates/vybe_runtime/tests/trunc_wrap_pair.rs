use vybe_runtime::{Chunk, Op, VM, Value};

fn convert(input: f64, separated: bool) -> Result<Value, String> {
    let mut chunk = Chunk::new("trunc_wrap_pair");
    chunk.emit_f64_const(input, 1);
    chunk.emit_op(Op::I64_TRUNC_F64_S, 2);
    if separated {
        chunk.emit_op(Op::NOP, 3);
    }
    chunk.emit_op(Op::I32_WRAP_I64, 4);
    VM::new().run(vec![chunk]).map_err(|error| error.to_string())
}

#[test]
fn trunc_wrap_accepts_i64_values_outside_i32_range() {
    for separated in [false, true] {
        for (input, expected) in [
            (0.0, 0), (-0.0, 0), (1.9, 1), (-1.9, -1),
            (2_147_483_648.0, i32::MIN), (4_294_967_295.0, -1),
            (4_294_967_296.0, 0), (-4_294_967_297.0, -1),
            (-9_223_372_036_854_775_808.0, 0),
            (9_223_372_036_854_774_784.0, -1024),
        ] {
            match convert(input, separated).expect("valid trunc/wrap pair trapped") {
                Value::I32(actual) => assert_eq!(actual, expected, "input={input:?}"),
                other => panic!("expected i32 result, got {other:?}"),
            }
        }
    }
}

#[test]
fn trunc_wrap_rejects_nan_infinity_and_i64_overflow() {
    for separated in [false, true] {
        for input in [
            f64::NAN, f64::INFINITY, f64::NEG_INFINITY,
            9_223_372_036_854_775_808.0, -9_223_372_036_854_777_856.0,
        ] {
            assert!(convert(input, separated).is_err(), "invalid input did not trap: {input:?}");
        }
    }
}

fn workload(separated: bool, rounds: i32) -> Chunk {
    let mut chunk = Chunk::new("trunc_wrap_loop");
    chunk.local_count = 3;
    chunk.emit_f64_const(4_294_967_295.0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);
    chunk.emit_loop_s(0);
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_op(Op::I64_TRUNC_F64_S, 0);
    if separated {
        chunk.emit_op(Op::NOP, 0);
    }
    chunk.emit_op(Op::I32_WRAP_I64, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 2, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_TEE, 1, 0);
    chunk.emit_i32_const(rounds, 0);
    chunk.emit_op(Op::I32_LT_U, 0);
    chunk.emit_br_if(0, 0);
    chunk.emit_end(0);
    chunk.emit_op_u16(Op::LOCAL_GET, 2, 0);
    chunk
}

#[test]
fn structured_conversion_loop_returns_wrapped_result() {
    for separated in [false, true] {
        match VM::new().run(vec![workload(separated, 32)]).unwrap() {
            Value::I32(actual) => assert_eq!(actual, -1),
            other => panic!("expected wrapped loop result, got {other:?}"),
        }
    }
}

#[test]
#[ignore = "paired VM conversion-loop timing, includes module loading"]
fn paired_conversion_loop_benchmark() {
    use std::{hint::black_box, time::Instant};
    let sample = |separated| {
        let chunks = vec![workload(separated, 200_000)];
        let mut vm = VM::new();
        let start = Instant::now();
        let result = vm.run(chunks).unwrap();
        let elapsed = start.elapsed().as_secs_f64();
        assert!(matches!(result, Value::I32(-1)));
        black_box(vm);
        elapsed
    };
    let mut ratios = Vec::new();
    for pair in 0..15 {
        let (adjacent, separated) = if pair % 2 == 0 {
            (sample(false), sample(true))
        } else {
            let separated = sample(true);
            (sample(false), separated)
        };
        if pair >= 2 {
            ratios.push(separated / adjacent);
        }
    }
    ratios.sort_by(f64::total_cmp);
    eprintln!("separator/adjacent median={:.3}x, range={:.3}..{:.3}",
        ratios[6], ratios[0], ratios[12]);
    eprintln!("The separator adds one NOP; this is not a fusion speedup measurement.");
}
