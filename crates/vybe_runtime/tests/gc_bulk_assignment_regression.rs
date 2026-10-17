//! Normalized-bytecode dispatch coverage, not a WASM loader conformance suite.
use std::sync::{Arc, Mutex};
use vybe_runtime::chunk::{Chunk, CompositeKind};
use vybe_runtime::opcode::Op;
use vybe_runtime::typedef::TypeDef;
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::vm::VM;

fn array_type(vm: &mut VM) -> usize {
    let mut definition = TypeDef::new("bulk-dispatch-array");
    definition.kind = CompositeKind::Array;
    vm.type_registry.register(definition)
}

fn array(type_id: usize, values: Vec<Value>) -> Value {
    let mut object = Object::new();
    object.type_id = type_id;
    object.kind = ObjectKind::Array(values);
    Value::Object(Arc::new(Mutex::new(object)))
}

fn contents(value: &Value) -> Vec<Value> {
    let Value::Object(object) = value else { panic!("expected object"); };
    let object = object.lock().unwrap();
    let ObjectKind::Array(values) = &object.kind else { panic!("expected array"); };
    values.clone()
}

fn global(chunk: &mut Chunk, name: &str) {
    let index = chunk.add_constant(Value::String(Arc::from(name)));
    chunk.emit_op_u32(Op::GLOBAL_GET, index, 1);
}

fn finish(chunk: &mut Chunk) {
    chunk.local_count = 1;
    chunk.scratch_high_water = 1;
    chunk.result_arity = 1;
    chunk.emit_i32_const(42, 1);
    chunk.emit_op(Op::RETURN, 1);
}

fn copy_program(dst: i32, src: i32, count: i32) -> Chunk {
    let mut chunk = Chunk::new("bulk-array-copy-dispatch");
    global(&mut chunk, "destination");
    chunk.emit_i32_const(dst, 1);
    global(&mut chunk, "source");
    chunk.emit_i32_const(src, 1);
    chunk.emit_i32_const(count, 1);
    chunk.emit_op(Op::ARRAY_COPY, 1);
    finish(&mut chunk);
    chunk
}

fn fill_program(start: i32, count: i32) -> Chunk {
    let mut chunk = Chunk::new("bulk-array-fill-dispatch");
    global(&mut chunk, "destination");
    chunk.emit_i32_const(start, 1);
    chunk.emit_i32_const(77, 1);
    chunk.emit_i32_const(count, 1);
    chunk.emit_op(Op::ARRAY_FILL, 1);
    finish(&mut chunk);
    chunk
}

#[test]
fn same_array_copy_executes_every_small_overlap_range() {
    for typed in [false, true] {
        let mut vm = VM::new();
        let type_id = if typed { array_type(&mut vm) } else { 0 };
        let seed: Vec<_> = (0..8).map(Value::I32).collect();
        for src in 0..=8usize {
            for dst in 0..=8usize {
                for count in 0..=8 - src.max(dst) {
                    let target = array(type_id, seed.clone());
                    vm.set_global("source", target.clone());
                    vm.set_global("destination", target.clone());
                    let snapshot = seed[src..src + count].to_vec();
                    let mut expected = seed.clone();
                    expected[dst..dst + count].clone_from_slice(&snapshot);
                    assert_eq!(vm.run(vec![copy_program(dst as i32, src as i32, count as i32)]).unwrap(),
                        Value::I32(42));
                    assert_eq!(contents(&target), expected,
                        "typed={typed} src={src} dst={dst} count={count}");
                }
            }
        }
    }
}

#[test]
fn gc_copy_invalid_ranges_trap_before_writing_for_same_and_distinct_arrays() {
    for same in [false, true] {
        let mut vm = VM::new();
        let type_id = array_type(&mut vm);
        for (dst, src, count) in [
            (-1, 0, 0), (0, -1, 0), (0, 0, -1),
            (9, 0, 0), (0, 9, 0), (7, 0, 2), (0, 7, 2),
            (i32::MAX, 0, i32::MAX),
        ] {
            let source = array(type_id, (0..8).map(Value::I32).collect());
            let destination = if same { source.clone() } else {
                array(type_id, (100..108).map(Value::I32).collect())
            };
            let before_src = contents(&source);
            let before_dst = contents(&destination);
            vm.set_global("source", source.clone());
            vm.set_global("destination", destination.clone());
            let error = vm.run(vec![copy_program(dst, src, count)]).unwrap_err();
            assert!(error.to_string().contains("out of bounds"), "{error}");
            assert_eq!(contents(&source), before_src);
            assert_eq!(contents(&destination), before_dst,
                "same={same} dst={dst} src={src} count={count}");
        }
    }
}

#[test]
fn distinct_array_copy_keeps_sources_and_copied_reference_identity() {
    let mut vm = VM::new();
    let type_id = array_type(&mut vm);
    let elements: Vec<_> = (0..8).map(|_| {
        Value::Object(Arc::new(Mutex::new(Object::new())))
    }).collect();
    let source = array(type_id, elements.clone());
    let destination = array(type_id, elements.iter().rev().cloned().collect());
    vm.set_global("source", source.clone());
    vm.set_global("destination", destination.clone());
    assert_eq!(vm.run(vec![copy_program(2, 1, 4)]).unwrap(), Value::I32(42));
    let mut expected: Vec<_> = elements.iter().rev().cloned().collect();
    expected[2..6].clone_from_slice(&elements[1..5]);
    for (actual, expected) in contents(&destination).iter().zip(&expected) {
        let (Value::Object(a), Value::Object(b)) = (actual, expected) else {
            panic!("expected object references");
        };
        assert!(Arc::ptr_eq(a, b));
    }
    for (actual, expected) in contents(&source).iter().zip(&elements) {
        let (Value::Object(a), Value::Object(b)) = (actual, expected) else {
            panic!("expected object references");
        };
        assert!(Arc::ptr_eq(a, b));
    }
    let before = contents(&destination);
    assert_eq!(vm.run(vec![copy_program(8, 8, 0)]).unwrap(), Value::I32(42));
    for (actual, expected) in contents(&destination).iter().zip(&before) {
        let (Value::Object(a), Value::Object(b)) = (actual, expected) else {
            panic!("expected object references");
        };
        assert!(Arc::ptr_eq(a, b));
    }
}

#[test]
fn gc_fill_traps_without_partial_writes_and_accepts_the_empty_end() {
    let mut vm = VM::new();
    let type_id = array_type(&mut vm);
    let seed: Vec<_> = (0..8).map(Value::I32).collect();
    for (start, count) in [(-1, 0), (0, -1), (9, 0), (7, 2), (i32::MAX, i32::MAX)] {
        let target = array(type_id, seed.clone());
        vm.set_global("destination", target.clone());
        let error = vm.run(vec![fill_program(start, count)]).unwrap_err();
        assert!(error.to_string().contains("out of bounds"), "{error}");
        assert_eq!(contents(&target), seed);
    }
    let target = array(type_id, seed.clone());
    vm.set_global("destination", target.clone());
    assert_eq!(vm.run(vec![fill_program(8, 0)]).unwrap(), Value::I32(42));
    assert_eq!(contents(&target), seed);
    assert_eq!(vm.run(vec![fill_program(2, 3)]).unwrap(), Value::I32(42));
    let mut expected = seed;
    expected[2..5].fill(Value::I32(77));
    assert_eq!(contents(&target), expected);
}

#[test]
fn null_gc_arrays_trap_even_for_zero_count() {
    let mut vm = VM::new();
    let type_id = array_type(&mut vm);
    let real = array(type_id, Vec::new());
    vm.set_global("destination", Value::TypedNull(type_id));
    assert!(vm.run(vec![fill_program(0, 0)]).unwrap_err().to_string().contains("null"));
    for null_source in [false, true] {
        vm.set_global("source", if null_source { Value::TypedNull(type_id) } else { real.clone() });
        vm.set_global("destination", if null_source { real.clone() } else { Value::TypedNull(type_id) });
        assert!(vm.run(vec![copy_program(0, 0, 0)]).unwrap_err().to_string().contains("null"));
    }
}

#[test]
fn dynamic_array_fill_and_copy_retain_clipping() {
    let mut vm = VM::new();
    let seed: Vec<_> = (0..8).map(Value::I32).collect();
    let target = array(0, seed.clone());
    vm.set_global("destination", target.clone());
    assert_eq!(vm.run(vec![fill_program(-1, 2)]).unwrap(), Value::I32(42));
    let mut expected = seed;
    expected[..2].fill(Value::I32(77));
    assert_eq!(contents(&target), expected);
    assert_eq!(vm.run(vec![fill_program(7, 4)]).unwrap(), Value::I32(42));
    expected[7] = Value::I32(77);
    assert_eq!(contents(&target), expected);
    vm.set_global("source", target.clone());
    assert_eq!(vm.run(vec![copy_program(0, -1, 2)]).unwrap(), Value::I32(42));
    assert_eq!(contents(&target), expected);
    assert_eq!(vm.run(vec![copy_program(0, 0, -1)]).unwrap(), Value::I32(42));
    assert_eq!(contents(&target), expected);
}
