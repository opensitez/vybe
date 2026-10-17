//! Shared-buffer vector fills must match the unfused store loop, including
//! its successful prefix before a bounds trap.
use vybe_runtime::{Chunk, Op, VM, Value};

const PAGE: usize = 65_536;
const PATTERN: [u8; 16] = [0, 1, 127, 128, 254, 255, 3, 17, 31, 63, 129, 191, 253, 7, 11, 19];

fn local(chunk: &mut Chunk, op: Op, slot: u16) {
    chunk.emit_op(op, 1);
    chunk.code.extend_from_slice(&slot.to_be_bytes());
}

fn program(dst: i32, len: u32, start: u32, force_scalar: bool) -> Chunk {
    let mut chunk = Chunk::new("shared vector fill");
    chunk.local_count = 4;
    chunk.result_arity = 1;
    for (slot, value) in [(0, dst), (1, len as i32), (2, start as i32)] {
        chunk.emit_i32_const(value, 1);
        local(&mut chunk, Op::LOCAL_SET, slot);
    }
    chunk.emit_op(Op::V128_CONST, 1);
    for byte in PATTERN { chunk.emit(byte, 1); }
    local(&mut chunk, Op::LOCAL_SET, 3);
    chunk.emit_op_u8_u8(Op::BLOCK, 0, 0, 1);
    chunk.emit_op_u8_u8(Op::LOOP, 0, 0, 1);
    local(&mut chunk, Op::LOCAL_GET, 2);
    local(&mut chunk, Op::LOCAL_GET, 1);
    chunk.emit_op(Op::I32_GE_U, 1);
    chunk.emit_op(Op::BR_IF, 1);
    chunk.emit_leb_u32(1, 1);
    local(&mut chunk, Op::LOCAL_GET, 0);
    local(&mut chunk, Op::LOCAL_GET, 2);
    chunk.emit_op(Op::I32_ADD, 1);
    local(&mut chunk, Op::LOCAL_GET, 3);
    chunk.emit_op(Op::V128_STORE, 1);
    if force_scalar { chunk.emit_op(Op::NOP, 1); }
    local(&mut chunk, Op::LOCAL_GET, 2);
    chunk.emit_i32_const(16, 1);
    chunk.emit_op(Op::I32_ADD, 1);
    local(&mut chunk, Op::LOCAL_SET, 2);
    chunk.emit_op(Op::BR, 1);
    chunk.emit_leb_u32(0, 1);
    chunk.emit_op(Op::END, 1);
    chunk.emit_op(Op::END, 1);
    local(&mut chunk, Op::LOCAL_GET, 2);
    chunk.emit_op(Op::RETURN, 1);
    chunk
}

fn execute(dst: i32, len: u32, start: u32, force_scalar: bool) -> (Option<i32>, Vec<u8>) {
    let mut vm = VM::new();
    vm.memory.with_buffer_mut(|memory| {
        memory.clear();
        memory.resize(PAGE, 0x5a);
    });
    let alias = vm.memory.clone();
    let result = match vm.run(vec![program(dst, len, start, force_scalar)]) {
        Ok(Value::I32(value)) => Some(value),
        Ok(other) => panic!("unexpected vector-fill result: {other:?}"),
        Err(error) => {
            assert!(error.to_string().contains("out of bounds"), "{error}");
            None
        }
    };
    let memory = vm.memory.read_array::<PAGE>(0).unwrap().to_vec();
    assert_eq!(alias.read_array::<PAGE>(0).unwrap().as_slice(), memory.as_slice());
    (result, memory)
}

#[test]
fn shared_vector_fill_matches_unfused_counter_and_every_byte() {
    for len in [0, 1, 15, 16, 17, 31, 32, 257] {
        for start in [0, 1, 7, 16, 33] {
            assert_eq!(execute(5, len, start, false), execute(5, len, start, true),
                "len={len} start={start}");
        }
    }
    let (counter, bytes) = execute(7, 17, 0, false);
    assert_eq!(counter, Some(32));
    assert!(bytes[..7].iter().all(|byte| *byte == 0x5a));
    assert_eq!(&bytes[7..23], &PATTERN);
    assert_eq!(&bytes[23..39], &PATTERN);
    assert!(bytes[39..].iter().all(|byte| *byte == 0x5a));
}

#[test]
fn shared_vector_fill_preserves_bounds_trap_prefix_and_wrapping_fallback() {
    for (dst, len, start) in [
        (65_520, 32, 0), (65_521, 1, 0), (65_520, 1, 0),
        (-1, 16, 0), (-16, 32, 16), (16, u32::MAX, 0xffff_fff0),
    ] {
        assert_eq!(execute(dst, len, start, false), execute(dst, len, start, true),
            "dst={dst} len={len} start={start}");
    }
    let (counter, bytes) = execute(65_520, 32, 0, false);
    assert_eq!(counter, None);
    assert!(bytes[..65_520].iter().all(|byte| *byte == 0x5a));
    assert_eq!(&bytes[65_520..], &PATTERN);
}
