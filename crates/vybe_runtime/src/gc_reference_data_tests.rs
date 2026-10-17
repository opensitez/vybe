//! Element-segment allocation must preserve native RTT and reference identity.
use crate::{VM, Value};
use crate::chunk::{Chunk, CompositeKind, TypeEntry};
use crate::opcode::Op;
use crate::value::{Object, ObjectKind};
use std::sync::Arc;

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

fn constructor(values: Vec<Value>, typeidx: u16, offset: i32, count: i32, dropped: bool) -> Chunk {
    let mut chunk = Chunk::new("script");
    chunk.types.push(TypeEntry {
        name: "reference_segment_array".into(),
        kind: CompositeKind::Array,
        fields: vec!["externref".into()],
        ..Default::default()
    });
    chunk.elem_segments.push(values);
    if dropped {
        operation(&mut chunk, Op::ELEM_DROP, &[]);
        chunk.code.push(0);
    }
    constant(&mut chunk, offset);
    constant(&mut chunk, count);
    operation(&mut chunk, Op::ARRAY_NEW_ELEM, &[typeidx, 0]);
    chunk
}

fn finish(mut chunk: Chunk) -> Chunk {
    operation(&mut chunk, Op::RETURN, &[]);
    chunk
}

fn self_copy(mut chunk: Chunk, dst: i32, src: i32, count: i32) -> Chunk {
    chunk.local_count = 1;
    operation(&mut chunk, Op::LOCAL_TEE, &[0]);
    constant(&mut chunk, dst);
    operation(&mut chunk, Op::LOCAL_GET, &[0]);
    constant(&mut chunk, src);
    constant(&mut chunk, count);
    operation(&mut chunk, Op::ARRAY_COPY, &[]);
    operation(&mut chunk, Op::LOCAL_GET, &[0]);
    finish(chunk)
}

#[test]
fn element_constructor_preserves_rtt_identity_and_reference_lifetime() {
    let text: Arc<str> = "retained element".into();
    let target = crate::heap::alloc(Object::new());
    let chunk = finish(constructor(vec![Value::String(text.clone()),
        Value::Object(target.clone()), Value::Null], 1, 0, 3, false));
    let mut vm = VM::new();
    let Value::Object(array) = vm.run(vec![chunk]).unwrap() else { panic!("expected array"); };
    {
        let array = array.lock().unwrap();
        assert!(vm.type_registry.get(array.type_id).is_some_and(|ty| ty.is_array()));
        let ObjectKind::Array(values) = &array.kind else { panic!("wrong backing"); };
        assert_eq!(values.len(), 3);
        let Value::String(value) = &values[0] else { panic!("wrong string variant"); };
        assert!(Arc::ptr_eq(value, &text));
        let Value::Object(value) = &values[1] else { panic!("wrong reference variant"); };
        assert!(Arc::ptr_eq(value, &target));
        assert!(matches!(values[2], Value::Null));
    }
    drop(vm);
    assert_eq!(Arc::strong_count(&text), 2);
    assert_eq!(Arc::strong_count(&target), 2);
    drop(array);
    assert_eq!(Arc::strong_count(&text), 1);
    assert_eq!(Arc::strong_count(&target), 1);
}

#[test]
fn element_constructor_activates_typed_get_bounds_without_changing_dynamic_arrays() {
    for index in [-1, 2] {
        let mut typed = constructor(vec![Value::Null; 2], 1, 0, 2, false);
        constant(&mut typed, index);
        operation(&mut typed, Op::ARRAY_GET, &[]);
        let error = VM::new().run(vec![finish(typed)]).unwrap_err();
        assert!(error.to_string().contains("out of bounds"), "{error}");
    }
    let mut dynamic = constructor(vec![Value::Null; 2], 0, 0, 2, false);
    constant(&mut dynamic, 2);
    operation(&mut dynamic, Op::ARRAY_GET, &[]);
    assert!(VM::new().run(vec![finish(dynamic)]).is_ok());
}

#[test]
fn reference_self_copy_keeps_memmove_order_and_shared_identity() {
    let first: Arc<str> = "first".into();
    let second: Arc<str> = "second".into();
    let third: Arc<str> = "third".into();
    let chunk = self_copy(constructor(vec![Value::String(first.clone()),
        Value::String(second.clone()), Value::String(third.clone())], 1, 0, 3, false), 1, 0, 2);
    let mut vm = VM::new();
    let Value::Object(array) = vm.run(vec![chunk]).unwrap() else { panic!("expected array"); };
    {
        let array = array.lock().unwrap();
        let ObjectKind::Array(values) = &array.kind else { panic!("wrong backing"); };
        for (value, expected) in values.iter().zip([&first, &first, &second]) {
            let Value::String(value) = value else { panic!("wrong reference variant"); };
            assert!(Arc::ptr_eq(value, expected));
        }
    }
    drop(vm);
    assert_eq!(Arc::strong_count(&first), 3);
    assert_eq!(Arc::strong_count(&second), 2);
    assert_eq!(Arc::strong_count(&third), 1);
    drop(array);
    assert_eq!(Arc::strong_count(&first), 1);
    assert_eq!(Arc::strong_count(&second), 1);
}

#[test]
fn native_reference_copy_bounds_trap_and_dynamic_copy_remains_lenient() {
    let typed = self_copy(constructor(vec![Value::Null; 3], 1, 0, 3, false), 0, 0, 4);
    let error = VM::new().run(vec![typed]).unwrap_err();
    assert!(error.to_string().contains("out of bounds"), "{error}");
    let dynamic = self_copy(constructor(vec![Value::Null; 3], 0, 0, 3, false), 0, 0, 4);
    assert!(VM::new().run(vec![dynamic]).is_ok());
}

#[test]
fn dropped_element_segment_zero_length_keeps_native_array_type() {
    let mut vm = VM::new();
    let Value::Object(array) = vm.run(vec![finish(constructor(vec![Value::Null],
        1, 0, 0, true))]).unwrap() else { panic!("expected array"); };
    let array = array.lock().unwrap();
    assert!(vm.type_registry.get(array.type_id).is_some_and(|ty| ty.is_array()));
    assert!(matches!(&array.kind, ObjectKind::Array(values) if values.is_empty()));
    for (offset, count) in [(1, 0), (0, 1)] {
        assert!(VM::new().run(vec![finish(constructor(vec![Value::Null],
            1, offset, count, true))]).is_err());
    }
}

#[test]
fn reference_segment_ranges_use_unsigned_operands_and_validate_before_allocation() {
    for (offset, count) in [(-1, 0), (0, -1), (1, 2), (3, 0)] {
        let error = VM::new().run(vec![finish(constructor(vec![Value::Null; 2],
            1, offset, count, false))]).unwrap_err();
        assert!(error.to_string().contains("out of bounds"), "{error}");
    }
}

fn local_reference_pair(alias: bool, write: bool) -> Chunk {
    let mut chunk = constructor(vec![Value::Null; 2], 1, 0, 1, false);
    chunk.local_count = 2;
    operation(&mut chunk, Op::LOCAL_SET, &[0]);
    constant(&mut chunk, 0);
    constant(&mut chunk, 2);
    operation(&mut chunk, Op::ARRAY_NEW_ELEM, &[1, 0]);
    operation(&mut chunk, Op::LOCAL_SET, &[1]);
    operation(&mut chunk, Op::LOCAL_GET, &[0]);
    operation(&mut chunk, Op::LOCAL_GET, &[if alias { 0 } else { 1 }]);
    if write {
        operation(&mut chunk, Op::LOCAL_SET, &[0]);
        operation(&mut chunk, Op::LOCAL_GET, &[0]);
    }
    // Arrays are eqrefs: unlike function or external references, ref.eq is
    // a standard Wasm identity check for these native typed allocations.
    operation(&mut chunk, Op::REF_EQ, &[]);
    finish(chunk)
}

#[test]
fn consecutive_local_reads_preserve_native_gc_reference_identity() {
    for history in [false, true] {
        for alias in [false, true] {
            let mut vm = VM::new();
            vm.record_error_context(history);
            assert_eq!(vm.run(vec![local_reference_pair(alias, false)]).unwrap(),
                Value::I32(i32::from(alias)));
        }
    }
}

#[test]
fn local_reference_read_before_write_preserves_retained_operand() {
    for history in [false, true] {
        for alias in [false, true] {
            let mut vm = VM::new();
            vm.record_error_context(history);
            assert_eq!(vm.run(vec![local_reference_pair(alias, true)]).unwrap(),
                Value::I32(i32::from(alias)));
        }
    }
}

#[test]
fn overwritten_local_does_not_release_a_retained_native_array() {
    for history in [false, true] {
        let first: Arc<str> = "retained through assignment".into();
        let second: Arc<str> = "released with overwritten local".into();
        let mut chunk = constructor(vec![Value::String(first.clone()),
            Value::String(second.clone())], 1, 0, 1, false);
        chunk.local_count = 2;
        operation(&mut chunk, Op::LOCAL_SET, &[0]);
        constant(&mut chunk, 0);
        constant(&mut chunk, 2);
        operation(&mut chunk, Op::ARRAY_NEW_ELEM, &[1, 0]);
        operation(&mut chunk, Op::LOCAL_SET, &[1]);
        operation(&mut chunk, Op::LOCAL_GET, &[0]);
        operation(&mut chunk, Op::LOCAL_GET, &[1]);
        operation(&mut chunk, Op::LOCAL_SET, &[0]);
        let mut vm = VM::new();
        vm.record_error_context(history);
        let Value::Object(array) = vm.run(vec![finish(chunk)]).unwrap() else {
            panic!("retained typed array was replaced");
        };
        {
            let array = array.lock().unwrap();
            assert!(vm.type_registry.get(array.type_id).is_some_and(|ty| ty.is_array()));
            let ObjectKind::Array(values) = &array.kind else { panic!("wrong backing"); };
            assert_eq!(values.len(), 1);
            let Value::String(value) = &values[0] else { panic!("wrong element type"); };
            assert!(Arc::ptr_eq(value, &first));
        }
        drop(vm);
        assert_eq!(Arc::strong_count(&first), 2);
        assert_eq!(Arc::strong_count(&second), 1);
        drop(array);
        assert_eq!(Arc::strong_count(&first), 1);
    }
}
