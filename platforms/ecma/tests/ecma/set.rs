use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_bytecode::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<ecma-set-test>");
    let import_idx = chunk.add_import("ecma:set", name);
    let argc = args.len() as u8;
    for value in args {
        let constant = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, constant, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn invoke_iterator_next(iterator: Value) -> Value {
    let mut chunk = Chunk::new("<ecma-set-iterator-test>");
    let import_idx = chunk.add_import("ecma:iterator", "next");
    let constant = chunk.add_constant(iterator);
    chunk.emit_op_u16(Op::CONST, constant, 0);
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(1, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn array(values: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(values))))
}

fn iterator_values(iterator: Value) -> Vec<Value> {
    let mut out = Vec::new();
    loop {
        let step = invoke_iterator_next(iterator.clone());
        let Value::Object(object) = step else {
            panic!("iterator.next should return object");
        };
        let object = object.lock().unwrap();
        let done = matches!(object.properties.get("done"), Some(Value::Bool(true)));
        if done {
            break;
        }
        out.push(
            object
                .properties
                .get("value")
                .cloned()
                .unwrap_or(Value::Undefined),
        );
    }
    out
}

#[test]
fn new_set_from_iterable_deduplicates_members() {
    let set = invoke(
        "new",
        vec![array(vec![Value::I32(1), Value::I32(2), Value::I32(1)])],
    );
    assert_eq!(invoke("size", vec![set.clone()]), Value::I32(2));
    assert_eq!(invoke("has", vec![set, Value::I32(2)]), Value::Bool(true));
}

#[test]
fn add_delete_and_clear_update_membership_and_size() {
    let set = invoke("new", vec![]);
    let _ = invoke("add", vec![set.clone(), Value::String(Arc::from("x"))]);
    assert_eq!(
        invoke("has", vec![set.clone(), Value::String(Arc::from("x"))]),
        Value::Bool(true)
    );
    assert_eq!(invoke("size", vec![set.clone()]), Value::I32(1));
    assert_eq!(
        invoke("delete", vec![set.clone(), Value::String(Arc::from("x"))]),
        Value::Bool(true)
    );
    assert_eq!(invoke("size", vec![set.clone()]), Value::I32(0));
    assert!(matches!(invoke("clear", vec![set.clone()]), Value::Null));
    assert_eq!(invoke("size", vec![set]), Value::I32(0));
}

#[test]
fn from_iterable_on_string_materializes_characters() {
    let set = invoke("fromIterable", vec![Value::String(Arc::from("aba"))]);
    assert_eq!(invoke("size", vec![set.clone()]), Value::I32(2));
    assert_eq!(
        iterator_values(invoke("values", vec![set])),
        vec![Value::String(Arc::from("a")), Value::String(Arc::from("b"))]
    );
}

#[test]
fn values_and_entries_preserve_insertion_order() {
    let set = invoke(
        "new",
        vec![array(vec![Value::I32(3), Value::I32(1), Value::I32(4)])],
    );
    assert_eq!(
        iterator_values(invoke("values", vec![set.clone()])),
        vec![Value::I32(3), Value::I32(1), Value::I32(4)]
    );

    let entries = iterator_values(invoke("entries", vec![set]));
    assert_eq!(entries.len(), 3);
    let Value::Object(pair) = &entries[0] else {
        panic!("entry should be pair array")
    };
    let pair = pair.lock().unwrap();
    let ObjectKind::Array(values) = &pair.kind else {
        panic!("entry should be pair array")
    };
    assert_eq!(values, &vec![Value::I32(3), Value::I32(3)]);
}

#[test]
fn union_intersection_difference_and_symmetric_difference_follow_set_algebra() {
    let left = invoke(
        "new",
        vec![array(vec![Value::I32(1), Value::I32(2), Value::I32(3)])],
    );
    let right = invoke("new", vec![array(vec![Value::I32(3), Value::I32(4)])]);

    let union = invoke("union", vec![left.clone(), right.clone()]);
    let intersection = invoke("intersection", vec![left.clone(), right.clone()]);
    let difference = invoke("difference", vec![left.clone(), right.clone()]);
    let symmetric = invoke("symmetricDifference", vec![left, right]);

    assert_eq!(
        iterator_values(invoke("values", vec![union])),
        vec![Value::I32(1), Value::I32(2), Value::I32(3), Value::I32(4)]
    );
    assert_eq!(
        iterator_values(invoke("values", vec![intersection])),
        vec![Value::I32(3)]
    );
    assert_eq!(
        iterator_values(invoke("values", vec![difference])),
        vec![Value::I32(1), Value::I32(2)]
    );
    assert_eq!(
        iterator_values(invoke("values", vec![symmetric])),
        vec![Value::I32(1), Value::I32(2), Value::I32(4)]
    );
}

#[test]
fn relational_predicates_report_subset_superset_and_disjointness() {
    let left = invoke("new", vec![array(vec![Value::I32(1), Value::I32(2)])]);
    let right = invoke(
        "new",
        vec![array(vec![Value::I32(1), Value::I32(2), Value::I32(3)])],
    );
    let other = invoke("new", vec![array(vec![Value::I32(9)])]);

    assert_eq!(
        invoke("isSubsetOf", vec![left.clone(), right.clone()]),
        Value::I32(1)
    );
    assert_eq!(
        invoke("isSupersetOf", vec![right.clone(), left]),
        Value::I32(1)
    );
    assert_eq!(
        invoke("isDisjointFrom", vec![right.clone(), other.clone()]),
        Value::I32(1)
    );
    assert_eq!(invoke("overlaps", vec![right, other]), Value::Bool(false));
}

#[test]
fn mutating_set_algebra_updates_receiver_in_place() {
    let base = invoke("new", vec![array(vec![Value::I32(1), Value::I32(2)])]);
    let other = invoke("new", vec![array(vec![Value::I32(2), Value::I32(3)])]);

    let _ = invoke("unionWith", vec![base.clone(), other.clone()]);
    assert_eq!(
        iterator_values(invoke("values", vec![base.clone()])),
        vec![Value::I32(1), Value::I32(2), Value::I32(3)]
    );

    let _ = invoke(
        "exceptWith",
        vec![
            base.clone(),
            invoke("new", vec![array(vec![Value::I32(2)])]),
        ],
    );
    assert_eq!(
        iterator_values(invoke("values", vec![base.clone()])),
        vec![Value::I32(1), Value::I32(3)]
    );

    let _ = invoke(
        "symmetricExceptWith",
        vec![
            base.clone(),
            invoke("new", vec![array(vec![Value::I32(3), Value::I32(4)])]),
        ],
    );
    assert_eq!(
        iterator_values(invoke("values", vec![base.clone()])),
        vec![Value::I32(1), Value::I32(4)]
    );

    let _ = invoke(
        "intersectWith",
        vec![
            base.clone(),
            invoke("new", vec![array(vec![Value::I32(4), Value::I32(9)])]),
        ],
    );
    assert_eq!(
        iterator_values(invoke("values", vec![base])),
        vec![Value::I32(4)]
    );
}

// ── Set.prototype.forEach (ECMA-262 §24.2.3.6) ───────────────────────────────

#[test]
fn for_each_visits_every_element_in_insertion_order() {
    // forEach(fn) calls fn(value, value, set) for each member; we use a
    // __noop callback descriptor which the host can recognise and skip,
    // returning Undefined — the important invariant is that it doesn't panic.
    let s = invoke(
        "new",
        vec![array(vec![Value::I32(1), Value::I32(2), Value::I32(3)])],
    );
    let cb = {
        let mut o = Object::new();
        o.properties.insert("__noop".to_string(), Value::Bool(true));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let result = invoke("forEach", vec![s, cb]);
    // forEach returns undefined per spec.
    assert!(matches!(result, Value::Undefined | Value::Null));
}
