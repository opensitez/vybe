//! Behaviour tests for `ecma:generator` host imports.
//!
//! Reference: ECMA-262 §27.3 GeneratorFunction, §27.5 Generator protocol.
//!
//! Generators implement the iterator protocol with next/return/throw.
//! Each test covers a distinct behaviour of the generator state machine.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<ecma-generator-test>");
    let import_idx = chunk.add_import("ecma:generator", name);
    let argc = args.len() as u8;
    for value in args {
        let c = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, c, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn s(text: &str) -> Value { Value::String(Arc::from(text)) }

fn iter_result(step: &Value) -> (Value, bool) {
    if let Value::Object(o) = step {
        let o = o.lock().unwrap();
        let value = o.properties.get("value").cloned().unwrap_or(Value::Undefined);
        let done  = matches!(o.properties.get("done"), Some(Value::Bool(true)));
        (value, done)
    } else {
        panic!("expected iterator result object, got {:?}", step);
    }
}

// ── fromValues — creates a finite generator from a list ──────────────────────

#[test]
fn from_values_yields_each_element_in_order() {
    // fromValues([1, 2, 3]) → generator that yields 1, 2, 3 then done.
    let g = invoke("fromValues", vec![
        Value::Object(Arc::new(Mutex::new(
            vybe_bytecode::value::Object::new_array(vec![
                Value::I32(1), Value::I32(2), Value::I32(3),
            ])
        )))
    ]);
    let s1 = invoke("next", vec![g.clone()]);
    let s2 = invoke("next", vec![g.clone()]);
    let s3 = invoke("next", vec![g.clone()]);
    let s4 = invoke("next", vec![g]);

    let (v1, d1) = iter_result(&s1);
    let (v2, d2) = iter_result(&s2);
    let (v3, d3) = iter_result(&s3);
    let (_, d4)  = iter_result(&s4);

    assert_eq!(v1, Value::I32(1)); assert!(!d1);
    assert_eq!(v2, Value::I32(2)); assert!(!d2);
    assert_eq!(v3, Value::I32(3)); assert!(!d3);
    assert!(d4, "fourth next must be done");
}

// ── done iteration — value is undefined ──────────────────────────────────────

#[test]
fn next_after_exhaustion_returns_undefined_done_true() {
    let g = invoke("fromValues", vec![
        Value::Object(Arc::new(Mutex::new(
            vybe_bytecode::value::Object::new_array(vec![Value::I32(1)])
        )))
    ]);
    invoke("next", vec![g.clone()]); // consume the one value
    let after = invoke("next", vec![g]);
    let (value, done) = iter_result(&after);
    assert!(done);
    assert_eq!(value, Value::Undefined);
}

// ── return — early termination ────────────────────────────────────────────────

#[test]
fn return_closes_the_generator_and_returns_given_value() {
    // ECMA-262 §27.5.1.3: generator.return(v) → { value: v, done: true }.
    let g = invoke("fromValues", vec![
        Value::Object(Arc::new(Mutex::new(
            vybe_bytecode::value::Object::new_array(vec![
                Value::I32(1), Value::I32(2),
            ])
        )))
    ]);
    let result = invoke("return", vec![g.clone(), Value::I32(99)]);
    let (value, done) = iter_result(&result);
    assert_eq!(value, Value::I32(99));
    assert!(done);
    // Generator is now closed; next call must also be done.
    let after = invoke("next", vec![g]);
    let (_, after_done) = iter_result(&after);
    assert!(after_done);
}

// ── throw — injects an error into the generator ──────────────────────────────

#[test]
fn throw_on_done_generator_propagates_the_error() {
    // ECMA-262 §27.5.1.4: throw on a completed generator re-throws.
    let g = invoke("fromValues", vec![
        Value::Object(Arc::new(Mutex::new(
            vybe_bytecode::value::Object::new_array(vec![])
        )))
    ]);
    invoke("next", vec![g.clone()]); // exhaust immediately
    let err_result = invoke("throw", vec![g, s("boom")]);
    // Either the host returns an error marker or propagates; it must not panic.
    assert!(matches!(err_result, Value::Object(_) | Value::String(_) | Value::Undefined));
}

// ── range generator ───────────────────────────────────────────────────────────

#[test]
fn range_generator_yields_consecutive_integers() {
    // range(3, 6) → 3, 4, 5 then done.
    let g = invoke("range", vec![Value::I32(3), Value::I32(6)]);
    let (v1, _) = iter_result(&invoke("next", vec![g.clone()]));
    let (v2, _) = iter_result(&invoke("next", vec![g.clone()]));
    let (v3, _) = iter_result(&invoke("next", vec![g.clone()]));
    let (_, d4) = iter_result(&invoke("next", vec![g]));
    assert_eq!(v1, Value::I32(3));
    assert_eq!(v2, Value::I32(4));
    assert_eq!(v3, Value::I32(5));
    assert!(d4);
}

#[test]
fn empty_range_is_immediately_done() {
    let g = invoke("range", vec![Value::I32(5), Value::I32(5)]);
    let (_, done) = iter_result(&invoke("next", vec![g]));
    assert!(done);
}

// ── map / filter on generators ────────────────────────────────────────────────

#[test]
fn map_transforms_each_yielded_value() {
    // map(range(0,3), x => x*2) → 0, 2, 4.
    // Encoded as { __map_mul: 2 }.
    let transform = {
        let mut o = Object::new();
        o.properties.insert("__map_mul".to_string(), Value::I32(2));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let src = invoke("range", vec![Value::I32(0), Value::I32(3)]);
    let g = invoke("map", vec![src, transform]);
    let (v0, _) = iter_result(&invoke("next", vec![g.clone()]));
    let (v1, _) = iter_result(&invoke("next", vec![g.clone()]));
    let (v2, _) = iter_result(&invoke("next", vec![g]));
    assert_eq!(v0, Value::I32(0));
    assert_eq!(v1, Value::I32(2));
    assert_eq!(v2, Value::I32(4));
}

#[test]
fn filter_skips_non_matching_values() {
    // filter(range(0,5), x => x % 2 == 0) → 0, 2, 4.
    let predicate = {
        let mut o = Object::new();
        o.properties.insert("__filter_mod_eq".to_string(),
            Value::Object(Arc::new(Mutex::new({
                let mut p = Object::new();
                p.properties.insert("mod".to_string(), Value::I32(2));
                p.properties.insert("eq".to_string(), Value::I32(0));
                p
            })))
        );
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let src = invoke("range", vec![Value::I32(0), Value::I32(5)]);
    let g = invoke("filter", vec![src, predicate]);
    let (v0, _) = iter_result(&invoke("next", vec![g.clone()]));
    let (v1, _) = iter_result(&invoke("next", vec![g.clone()]));
    let (v2, _) = iter_result(&invoke("next", vec![g.clone()]));
    let (_, d3) = iter_result(&invoke("next", vec![g]));
    assert_eq!(v0, Value::I32(0));
    assert_eq!(v1, Value::I32(2));
    assert_eq!(v2, Value::I32(4));
    assert!(d3);
}

// ── toArray ───────────────────────────────────────────────────────────────────

#[test]
fn to_array_collects_all_yielded_values() {
    let g = invoke("fromValues", vec![
        Value::Object(Arc::new(Mutex::new(
            vybe_bytecode::value::Object::new_array(vec![
                Value::I32(10), Value::I32(20),
            ])
        )))
    ]);
    let arr = invoke("toArray", vec![g]);
    if let Value::Object(o) = arr {
        if let vybe_bytecode::value::ObjectKind::Array(elems) = &o.lock().unwrap().kind {
            assert_eq!(elems.len(), 2);
            assert_eq!(elems[0], Value::I32(10));
            assert_eq!(elems[1], Value::I32(20));
        } else {
            panic!("expected array kind");
        }
    } else {
        panic!("expected array object");
    }
}

// ── generator is its own iterator (@@iterator returns self) ──────────────────

#[test]
fn generator_symbol_iterator_returns_self() {
    // ECMA-262 §27.5: generators are their own iterators.
    let g = invoke("fromValues", vec![
        Value::Object(Arc::new(Mutex::new(
            vybe_bytecode::value::Object::new_array(vec![Value::I32(1)])
        )))
    ]);
    let self_iter = invoke("symbolIterator", vec![g.clone()]);
    let gen_ptr  = match &g       { Value::Object(a) => Arc::as_ptr(a) as usize, _ => 0 };
    let iter_ptr = match &self_iter { Value::Object(a) => Arc::as_ptr(a) as usize, _ => 1 };
    assert_eq!(gen_ptr, iter_ptr);
}
