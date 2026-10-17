use super::VM;
use crate::{chunk::{Chunk, CompositeKind, TypeEntry}, opcode::Op, value::{ObjectKind, Value}};

fn op(chunk: &mut Chunk, operation: Op) {
    chunk.code.extend_from_slice(&operation.encode());
}

fn integer(chunk: &mut Chunk, mut value: i32) {
    op(chunk, Op::I32_CONST);
    loop {
        let byte = (value & 127) as u8;
        value >>= 7;
        let done = (value == 0 && byte & 64 == 0)
            || (value == -1 && byte & 64 != 0);
        chunk.code.push(if done { byte } else { byte | 128 });
        if done { break; }
    }
}

fn program(storage: &str, typed: bool, count: i32, default: bool) -> Chunk {
    let mut chunk = Chunk::new("script");
    if typed {
        chunk.types.push(TypeEntry {
            name: format!("gc_array_{storage}"),
            kind: CompositeKind::Array,
            fields: vec![storage.into()],
            ..Default::default()
        });
    }
    if !default { integer(&mut chunk, -37); }
    integer(&mut chunk, count);
    op(&mut chunk, if default { Op::ARRAY_NEW_DEFAULT } else { Op::ARRAY_NEW });
    chunk.code.extend_from_slice(&(if typed { 1u16 } else { 0u16 }).to_be_bytes());
    op(&mut chunk, Op::RETURN);
    chunk
}

#[test]
fn array_new_default_uses_declared_storage_defaults_and_native_rtt() {
    for (storage, expected) in [
        ("i8", Value::I32(0)), ("i16", Value::I32(0)), ("i32", Value::I32(0)),
        ("i64", Value::I64(0)), ("f32", Value::F32(0.0)), ("f64", Value::F64(0.0)),
        ("v128", Value::V128([0; 16])), ("externref", Value::Null),
    ] {
        let mut vm = VM::new();
        let result = vm.run(vec![program(storage, true, 17, true)]).unwrap();
        let Value::Object(object) = result else { panic!("constructor did not return an object"); };
        let object = object.lock().unwrap();
        assert!(object.type_id > 0);
        assert!(vm.type_registry.get(object.type_id).is_some_and(|ty| ty.is_array()));
        let ObjectKind::Array(values) = &object.kind else { panic!("constructor did not return array storage"); };
        assert_eq!(values.len(), 17);
        for value in values {
            assert_eq!(std::mem::discriminant(value), std::mem::discriminant(&expected));
            assert_eq!(value, &expected);
        }
    }
}

#[test]
fn typed_array_new_lengths_are_unsigned_and_exhaustion_traps() {
    for default in [false, true] {
        for count in [-1, -2, crate::vm::TABLE_ALLOC_LIMIT as i32 + 1] {
            let mut vm = VM::new();
            let error = vm.run(vec![program("i32", true, count, default)]).unwrap_err();
            assert!(error.to_string().contains("allocation failure"), "{error}");
        }
    }
}

#[test]
fn dynamic_array_negative_lengths_keep_existing_empty_array_behavior() {
    for default in [false, true] {
        let mut vm = VM::new();
        let result = vm.run(vec![program("i32", false, -1, default)]).unwrap();
        let Value::Object(object) = result else { panic!("constructor did not return an object"); };
        let object = object.lock().unwrap();
        let ObjectKind::Array(values) = &object.kind else { panic!("constructor did not return array storage"); };
        assert!(values.is_empty());
    }
}

#[test]
fn typed_array_new_preserves_supplied_integer_values() {
    let mut vm = VM::new();
    let result = vm.run(vec![program("i32", true, 257, false)]).unwrap();
    let Value::Object(object) = result else { panic!("constructor did not return an object"); };
    let object = object.lock().unwrap();
    let ObjectKind::Array(values) = &object.kind else { panic!("constructor did not return array storage"); };
    assert_eq!(values, &vec![Value::I32(-37); 257]);
}
