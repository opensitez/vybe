//! Aliased SIMD copies must preserve sequential overlap and trap-prefix effects.
use vybe_runtime::{Chunk, Op, VM, Value};

const PAGE: usize = 65_536;
const INITIAL: [u8; 16] = [0x73; 16];
type Case = (i32, i32, u32, u32, u32, u32);

#[derive(Debug, PartialEq)]
enum Observed { Counter(i32), Vector([u8; 16]), Trap }

fn local(chunk: &mut Chunk, op: Op, slot: u16) {
    chunk.emit_op(op, 1);
    chunk.code.extend_from_slice(&slot.to_be_bytes());
}

fn memory_op(chunk: &mut Chunk, op: Op, offset: u32) {
    chunk.emit_op(op, 1);
    if offset != 0 {
        chunk.emit_leb_u32(0x80, 1);
        chunk.emit_leb_u32(offset, 1);
    }
}

fn program(case: Case, force_scalar: bool, return_counter: bool) -> Chunk {
    let (src, dst, len, start, load_offset, store_offset) = case;
    let mut chunk = Chunk::new("shared vector copy");
    chunk.local_count = 5;
    chunk.result_arity = 1;
    for (slot, value) in [(0, src), (1, dst), (2, len as i32), (3, start as i32)] {
        chunk.emit_i32_const(value, 1);
        local(&mut chunk, Op::LOCAL_SET, slot);
    }
    chunk.emit_op(Op::V128_CONST, 1);
    for byte in INITIAL { chunk.emit(byte, 1); }
    local(&mut chunk, Op::LOCAL_SET, 4);
    chunk.emit_op_u8_u8(Op::BLOCK, 0, 0, 1);
    chunk.emit_op_u8_u8(Op::LOOP, 0, 0, 1);
    local(&mut chunk, Op::LOCAL_GET, 3);
    local(&mut chunk, Op::LOCAL_GET, 2);
    chunk.emit_op(Op::I32_GE_U, 1);
    chunk.emit_op(Op::BR_IF, 1);
    chunk.emit_leb_u32(1, 1);
    local(&mut chunk, Op::LOCAL_GET, 0);
    local(&mut chunk, Op::LOCAL_GET, 3);
    chunk.emit_op(Op::I32_ADD, 1);
    memory_op(&mut chunk, Op::V128_LOAD, load_offset);
    local(&mut chunk, Op::LOCAL_SET, 4);
    local(&mut chunk, Op::LOCAL_GET, 1);
    local(&mut chunk, Op::LOCAL_GET, 3);
    chunk.emit_op(Op::I32_ADD, 1);
    local(&mut chunk, Op::LOCAL_GET, 4);
    memory_op(&mut chunk, Op::V128_STORE, store_offset);
    if force_scalar { chunk.emit_op(Op::NOP, 1); }
    local(&mut chunk, Op::LOCAL_GET, 3);
    chunk.emit_i32_const(16, 1);
    chunk.emit_op(Op::I32_ADD, 1);
    local(&mut chunk, Op::LOCAL_SET, 3);
    chunk.emit_op(Op::BR, 1);
    chunk.emit_leb_u32(0, 1);
    chunk.emit_op(Op::END, 1);
    chunk.emit_op(Op::END, 1);
    local(&mut chunk, Op::LOCAL_GET, if return_counter { 3 } else { 4 });
    chunk.emit_op(Op::RETURN, 1);
    chunk
}

fn execute(case: Case, force_scalar: bool, return_counter: bool) -> (Observed, Vec<u8>) {
    let mut vm = VM::new();
    vm.memory.with_buffer_mut(|memory| {
        memory.clear();
        memory.extend((0..PAGE).map(|index| (index.wrapping_mul(37) + 11) as u8));
    });
    let alias = vm.memory.clone();
    let result = match vm.run(vec![program(case, force_scalar, return_counter)]) {
        Ok(Value::I32(value)) => Observed::Counter(value),
        Ok(Value::V128(bytes)) => Observed::Vector(bytes),
        Ok(other) => panic!("unexpected copy result: {other:?}"),
        Err(error) => {
            assert!(error.to_string().contains("out of bounds"), "{error}");
            Observed::Trap
        }
    };
    let bytes = vm.memory.read_array::<PAGE>(0).unwrap().to_vec();
    assert_eq!(alias.read_array::<PAGE>(0).unwrap().as_slice(), bytes.as_slice());
    (result, bytes)
}

fn compare(case: Case) {
    for return_counter in [false, true] {
        assert_eq!(execute(case, false, return_counter), execute(case, true, return_counter),
            "case={case:?} return_counter={return_counter}");
    }
}

#[test]
fn shared_vector_copy_preserves_bytes_last_value_and_counter() {
    for len in [0, 1, 15, 16, 17, 31, 32, 257] {
        for start in [0, 7, 16, 33] {
            for (src, dst) in [(5, 2048), (5, 5), (5, 8), (8, 5)] {
                compare((src, dst, len, start, 0, 0));
            }
        }
    }
    compare((0, 2048, 33, 0, 7, 11));
    compare((4, 4, 48, 0, 7, 3));
    let (last, bytes) = execute((5, 2048, 17, 0, 0, 0), false, false);
    let expected: [u8; 16] = std::array::from_fn(|lane| ((21 + lane) * 37 + 11) as u8);
    assert_eq!(last, Observed::Vector(expected));
    assert_eq!(&bytes[2064..2080], &expected);
}

#[test]
fn shared_vector_copy_preserves_trap_prefix_and_displacement_overflow() {
    for case in [
        (65_520, 32_768, 32, 0, 0, 0), (32_768, 65_520, 32, 0, 0, 0),
        (-1, 2048, 16, 0, 0, 0), (-16, 2048, 32, 16, 0, 0),
        (16, 2048, 16, 0, u32::MAX, 0), (16, 2048, 16, 0, 0, u32::MAX),
        (16, 32, u32::MAX, 0xffff_fff0, 0, 0),
    ] { compare(case); }
    let (result, bytes) = execute((65_520, 32_768, 32, 0, 0, 0), false, false);
    assert_eq!(result, Observed::Trap);
    let expected: [u8; 16] = std::array::from_fn(|lane| ((65_520 + lane) * 37 + 11) as u8);
    assert_eq!(&bytes[32_768..32_784], &expected);
}
