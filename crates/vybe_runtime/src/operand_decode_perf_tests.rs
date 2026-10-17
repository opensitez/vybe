use super::*;
use std::hint::black_box;
use std::time::Instant;

fn memory_candidate(op: Op) -> bool {
    (Op::I32_LOAD.0..=Op::I64_STORE32.0).contains(&op.0)
        || (op.group() == 0xFD
            && op.operand_format() == crate::opcode::OperandFormat::SimdMemArg)
}

#[test]
fn cheap_memory_classification_matches_every_valid_opcode_format() {
    for group in [0x00, 0xF0, 0xF1, 0xFB, 0xFC, 0xFD, 0xFE, 0xFF] {
        for sub in 0..=u16::MAX {
            if let Some(op) = Op::decode(group, sub) {
                assert_eq!(memory_candidate(op),
                    op.operand_format() == crate::opcode::OperandFormat::SimdMemArg,
                    "memory operand classification differs for {op:?}");
            }
        }
    }
}

#[test]
#[ignore = "mixed operand decoder helper; excludes VM execution and allocation"]
fn operand_decode_classification_loaded_benchmark() {
    let cases: Vec<(Op, Vec<u8>)> = vec![
        (Op::I32_CONST, vec![42]),
        (Op::I32_ADD, vec![]),
        (Op::MEMORY_COPY, vec![0, 0]),
        (Op::V128_LOAD, vec![0x84, 1, 16]),
        (Op::LOCAL_GET, vec![0, 1]),
        (Op::GLOBAL_GET, vec![0, 0, 0, 7]),
        (Op::F32_CONST, vec![1, 2, 3, 4]),
        (Op::DROP, vec![]),
        (Op::DATA_DROP, vec![0]),
        (Op::STRUCT_GET, vec![0, 0]),
    ];
    let expected: u64 = cases.iter().map(|(op, code)| {
        VM::decode_operand32(code, *op, 0) as u64
    }).sum::<u64>() * 100_000;
    let mut samples = Vec::new();
    for round in 0..23 {
        let start = Instant::now();
        let mut checksum = 0u64;
        for iteration in 0..1_000_000usize {
            let (op, code) = &cases[iteration % cases.len()];
            checksum += VM::decode_operand32(black_box(code), black_box(*op), 0) as u64;
        }
        let elapsed = start.elapsed().as_secs_f64();
        assert_eq!(black_box(checksum), expected);
        if round >= 2 { samples.push(elapsed); }
    }
    samples.sort_by(f64::total_cmp);
    eprintln!("mixed-operand-decoder p50_ms={:.3} p95_ms={:.3} checksum={expected}",
        samples[10] * 1000.0, samples[19] * 1000.0);
}
