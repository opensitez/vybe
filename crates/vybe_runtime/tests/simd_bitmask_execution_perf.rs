use std::hint::black_box;
use std::time::Instant;
use vybe_runtime::{Chunk, Op, VM, Value};

fn emit_vector(chunk: &mut Chunk, bytes: &[u8; 16]) {
    chunk.emit_op(Op::V128_CONST, 0);
    for byte in bytes { chunk.emit(*byte, 0); }
}

fn expected(bytes: &[u8; 16]) -> i32 {
    bytes.iter().enumerate().fold(0, |mask, (lane, byte)| {
        mask | (i32::from((*byte as i8).is_negative()) << lane)
    })
}

#[test]
fn bitmask_dispatch_preserves_lane_order_and_result_type() {
    let mut masks = vec![0u16, u16::MAX, 0xaaaa, 0x5555];
    for lane in 0..16 {
        masks.push(1u16 << lane);
        masks.push(!(1u16 << lane));
    }
    for mask in masks {
        let mut bytes = [0u8; 16];
        for (lane, byte) in bytes.iter_mut().enumerate() {
            *byte = (((mask >> lane) & 1) as u8) << 7 | ((lane * 13) & 127) as u8;
        }
        let mut chunk = Chunk::new("<script>");
        emit_vector(&mut chunk, &bytes);
        chunk.emit_op(Op::I8X16_BITMASK, 0);
        chunk.emit_op(Op::RETURN, 0);
        let result = VM::new().run(vec![chunk]).unwrap();
        assert!(matches!(result, Value::I32(value) if value == expected(&bytes)));
    }
}

#[test]
fn bitmask_dispatch_replaces_exactly_one_stack_operand() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_i32_const(7654321, 0);
    emit_vector(&mut chunk, &[0x80; 16]);
    chunk.emit_op(Op::I8X16_BITMASK, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::RETURN, 0);
    assert!(matches!(VM::new().run(vec![chunk]).unwrap(), Value::I32(7654321)));
}

#[test]
#[ignore = "loaded prepared-bytecode module run, including setup/decode"]
fn bitmask_module_run_loaded_benchmark() {
    const COUNT: usize = 100_000;
    let bytes = [0, 255, 127, 128, 91, 171, 3, 193, 239, 2, 44, 205, 78, 148, 113, 129];
    let mut times = Vec::with_capacity(21);
    for _ in 0..21 {
        let mut chunk = Chunk::new("<script>");
        for iteration in 0..COUNT {
            emit_vector(&mut chunk, &bytes);
            chunk.emit_op(Op::I8X16_BITMASK, 0);
            if iteration + 1 < COUNT { chunk.emit_op(Op::DROP, 0); }
        }
        chunk.emit_op(Op::RETURN, 0);
        let mut vm = VM::new();
        let start = Instant::now();
        let result = vm.run(vec![chunk]).unwrap();
        let elapsed = start.elapsed().as_nanos();
        assert!(matches!(result, Value::I32(value) if value == expected(&bytes)));
        black_box(result);
        times.push(elapsed);
    }
    times.sort_unstable();
    println!("bitmask module count={COUNT} p50_ns={} p95_ns={}; includes VM::run setup/decode", times[10], times[19]);
}
