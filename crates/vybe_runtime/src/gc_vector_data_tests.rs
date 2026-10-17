//! GC data-array regressions: storage widths, exact bits, and segment bounds.
use crate::{VM, Value};
use crate::chunk::{Chunk, CompositeKind, TypeEntry};
use crate::opcode::Op;
use crate::value::ObjectKind;

fn constant(chunk: &mut Chunk, mut value: i32) {
    chunk.code.extend_from_slice(&Op::I32_CONST.encode());
    loop {
        let byte = (value & 127) as u8;
        value >>= 7;
        let done = (value == 0 && byte & 64 == 0) || (value == -1 && byte & 64 != 0);
        chunk.code.push(if done { byte } else { byte | 128 });
        if done { break; }
    }
}

fn operation(chunk: &mut Chunk, op: Op, operands: &[u16]) {
    chunk.code.extend_from_slice(&op.encode());
    for operand in operands { chunk.code.extend_from_slice(&operand.to_be_bytes()); }
}

fn module(storage: &str, data: Vec<u8>) -> Chunk {
    let mut chunk = Chunk::new("script");
    chunk.types.push(TypeEntry {
        name: format!("data_array_{storage}"),
        kind: CompositeKind::Array,
        fields: vec![storage.into()],
        ..Default::default()
    });
    chunk.data_segments.push(data);
    chunk
}

fn new_data(storage: &str, data: Vec<u8>, offset: i32, count: i32, dropped: bool) -> Chunk {
    let mut chunk = module(storage, data);
    if dropped {
        operation(&mut chunk, Op::DATA_DROP, &[]);
        chunk.code.push(0); // dataidx: unsigned LEB.
    }
    constant(&mut chunk, offset);
    constant(&mut chunk, count);
    operation(&mut chunk, Op::ARRAY_NEW_DATA, &[1, 0]);
    operation(&mut chunk, Op::RETURN, &[]);
    chunk
}

fn init_data(data: Vec<u8>, dst: i32, src: i32, count: i32, dropped: bool) -> Chunk {
    let mut chunk = module("v128", data);
    chunk.local_count = 1;
    if dropped {
        operation(&mut chunk, Op::DATA_DROP, &[]);
        chunk.code.push(0);
    }
    constant(&mut chunk, 4);
    operation(&mut chunk, Op::ARRAY_NEW_DEFAULT, &[1]);
    operation(&mut chunk, Op::LOCAL_TEE, &[0]);
    constant(&mut chunk, dst);
    constant(&mut chunk, src);
    constant(&mut chunk, count);
    operation(&mut chunk, Op::ARRAY_INIT_DATA, &[1, 0]);
    operation(&mut chunk, Op::LOCAL_GET, &[0]);
    operation(&mut chunk, Op::RETURN, &[]);
    chunk
}

fn array_values(chunk: Chunk) -> Vec<Value> {
    let mut vm = VM::new();
    let result = vm.run(vec![chunk]).unwrap();
    let Value::Object(object) = result else { panic!("expected array"); };
    let object = object.lock().unwrap();
    assert!(vm.type_registry.get(object.type_id).is_some_and(|ty| ty.is_array()));
    let ObjectKind::Array(values) = &object.kind else { panic!("wrong array backing"); };
    values.clone()
}

#[test]
fn new_data_vectors_preserve_unaligned_bytes_and_native_rtt() {
    let data: Vec<_> = (0u8..35).map(|byte| byte.wrapping_mul(37)).collect();
    let expected = vec![
        Value::V128(data[3..19].try_into().unwrap()),
        Value::V128(data[19..35].try_into().unwrap()),
    ];
    assert_eq!(array_values(new_data("v128", data, 3, 2, false)), expected);
}

#[test]
fn init_data_vectors_preserve_untouched_lanes_and_exact_bytes() {
    let data: Vec<_> = (0u8..35).map(|byte| byte.wrapping_mul(37)).collect();
    let expected = vec![
        Value::V128([0; 16]),
        Value::V128(data[3..19].try_into().unwrap()),
        Value::V128(data[19..35].try_into().unwrap()),
        Value::V128([0; 16]),
    ];
    assert_eq!(array_values(init_data(data, 1, 3, 2, false)), expected);
}

#[test]
fn vector_data_bounds_count_sixteen_bytes_per_element() {
    for chunk in [
        new_data("v128", vec![0; 15], 0, 1, false),
        new_data("v128", vec![0; 31], 0, 2, false),
        init_data(vec![0; 15], 0, 0, 1, false),
        init_data(vec![0; 32], 3, 0, 2, false),
    ] {
        let error = VM::new().run(vec![chunk]).unwrap_err();
        assert!(error.to_string().contains("out of bounds"), "{error}");
    }
}

#[test]
fn dropped_vector_segments_are_empty_not_unconditionally_invalid() {
    assert!(array_values(new_data("v128", vec![0; 16], 0, 0, true)).is_empty());
    assert_eq!(array_values(init_data(vec![0; 16], 4, 0, 0, true)),
        vec![Value::V128([0; 16]); 4]);
    for chunk in [
        new_data("v128", vec![0; 16], 1, 0, true),
        new_data("v128", vec![0; 16], 0, 1, true),
        init_data(vec![0; 16], 0, 1, 0, true),
        init_data(vec![0; 16], 0, 0, 1, true),
    ] {
        assert!(VM::new().run(vec![chunk]).is_err());
    }
}

#[test]
fn data_float_decoding_preserves_zero_and_nan_payload_bits() {
    let f32_bits = [0x8000_0000u32, 0x7fc1_2345, 0x7fa1_2345, 0x3f80_0000];
    let f64_bits = [0x8000_0000_0000_0000u64, 0x7ff8_1234_5678_9abc,
        0x7ff0_0000_0000_1234, 0x3ff0_0000_0000_0000];
    let mut data = vec![0; 3];
    data.extend(f32_bits.iter().flat_map(|bits| bits.to_le_bytes()));
    for (value, bits) in array_values(new_data("f32", data, 3, 4, false)).iter().zip(f32_bits) {
        let Value::F32(value) = value else { panic!("wrong f32 variant"); };
        assert_eq!(value.to_bits(), bits);
    }
    let mut data = vec![0; 3];
    data.extend(f64_bits.iter().flat_map(|bits| bits.to_le_bytes()));
    for (value, bits) in array_values(new_data("f64", data, 3, 4, false)).iter().zip(f64_bits) {
        let Value::F64(value) = value else { panic!("wrong f64 variant"); };
        assert_eq!(value.to_bits(), bits);
    }
}

#[test]
fn data_array_reservation_rejects_resource_exhaustion_before_initialization() {
    for len in [crate::vm::TABLE_ALLOC_LIMIT as usize + 1, usize::MAX] {
        assert!(super::reserve_array_values(len).is_err());
    }
    assert!(super::reserve_array_values(0).unwrap().is_empty());
}
