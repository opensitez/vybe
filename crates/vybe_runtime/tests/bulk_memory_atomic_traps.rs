use vybe_runtime::{Chunk, Value, VM, opcode::Op};

fn execute(fill: bool, destination: u32, source_or_value: u32, count: u32)
    -> (bool, Vec<u8>)
{
    let mut vm = VM::new();
    vm.memory.with_buffer_mut(|memory| {
        *memory = (0..64).map(|i| (i * 37 + 11) as u8).collect();
    });
    let mut chunk = Chunk::new("bulk-memory-atomic-trap");
    for value in [destination, source_or_value, count] {
        chunk.emit_i32_const(value as i32, 0);
    }
    if fill {
        chunk.emit_op_idx(Op::MEMORY_FILL, 0u32, 0);
    } else {
        chunk.emit_op_idx_idx(Op::MEMORY_COPY, 0u32, 0u32, 0);
    }
    chunk.emit_i32_const(42, 0);
    chunk.emit_op(Op::RETURN, 0);
    let result = vm.run(vec![chunk]);
    if let Ok(value) = &result { assert_eq!(value, &Value::I32(42)); }
    (result.is_err(), vm.memory.with_buffer(|memory| memory.to_vec()))
}

#[test]
fn copy_checks_both_full_ranges_before_any_write() {
    let original: Vec<u8> = (0..64).map(|i| (i * 37 + 11) as u8).collect();
    for (destination, source, count) in [
        (0u32, 0u32, 64u32), (1, 0, 63), (0, 1, 63),
        (8, 0, 16), (0, 8, 16), (16, 16, 16),
        (64, 64, 0), (64, 0, 0), (0, 64, 0),
        (63, 0, 2), (0, 63, 2), (63, 63, 2),
        (65, 0, 0), (0, 65, 0), (u32::MAX, 0, 0),
        (0, u32::MAX, 0), (u32::MAX, 0, 2),
        (0, u32::MAX, 2), (0, 0, u32::MAX),
    ] {
        let valid = u64::from(destination) + u64::from(count) <= 64
            && u64::from(source) + u64::from(count) <= 64;
        let mut expected = original.clone();
        if valid {
            expected.copy_within(source as usize..(source + count) as usize,
                destination as usize);
        }
        let (trapped, actual) = execute(false, destination, source, count);
        assert_eq!(trapped, !valid, "dst={destination} src={source} count={count}");
        assert_eq!(actual, expected, "dst={destination} src={source} count={count}");
    }
}

#[test]
fn fill_checks_full_range_and_uses_only_low_value_byte() {
    let original: Vec<u8> = (0..64).map(|i| (i * 37 + 11) as u8).collect();
    for value in [0u32, 255, 256, 0x1234abcd, u32::MAX] {
        for (destination, count) in [
            (0u32, 64u32), (3, 17), (63, 1), (64, 0),
            (63, 2), (65, 0), (u32::MAX, 0), (u32::MAX, 2), (0, u32::MAX),
        ] {
            let valid = u64::from(destination) + u64::from(count) <= 64;
            let mut expected = original.clone();
            if valid { expected[destination as usize..(destination + count) as usize].fill(value as u8); }
            let (trapped, actual) = execute(true, destination, value, count);
            assert_eq!(trapped, !valid, "dst={destination} value={value} count={count}");
            assert_eq!(actual, expected, "dst={destination} value={value} count={count}");
        }
    }
}
