use std::sync::{Arc, Mutex, Weak};
use vybe_compiler::primitives::{functions, globals, platforms, references};
use vybe_runtime::{Chunk, Op, VM, Value, value::Object};
fn guard(chunks: &mut [Chunk], old: bool) {
    if old {
        references::emit_autoderef_to_stack(chunks, 1, 37)
    } else {
        references::emit_nonreference_temporary_to_stack(chunks, 1, 37)
    }
}
fn module(old: bool, prior: bool) -> Vec<Chunk> {
    let mut entry = Chunk::new("<scalar-owner>");
    functions::emit_ref_func(&mut entry, 1, 0, 37);
    entry.emit_op(Op::RETURN, 37);
    let mut f = Chunk::new("scalar_producer_consumer");
    f.arity = 1;
    f.param_count = 1;
    f.local_count = 1;
    let mut chunks = vec![entry, f];
    if prior {
        let make = chunks[1].add_import("test:scalar", "makePrior");
        chunks[1].emit_call(make, 0, 37);
        references::emit_autoderef_to_stack(&mut chunks, 1, 37);
        chunks[1].emit_op(Op::DROP, 37);
    }
    chunks[1].emit_op_u16(Op::LOCAL_GET, 0, 37);
    guard(&mut chunks, old);
    if prior {
        chunks[1].emit_op(Op::DROP, 37);
        let observe = chunks[1].add_import("test:scalar", "observePrior");
        chunks[1].emit_call(observe, 0, 37);
    }
    chunks[1].emit_op(Op::RETURN, 37);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
fn exact(a: Value, b: Value) {
    match (&a, &b) {
        (Value::F64(a), Value::F64(b)) => assert_eq!(a.to_bits(), b.to_bits()),
        (Value::String(a), Value::String(b)) => assert!(Arc::ptr_eq(a, b)),
        _ => assert_eq!(a, b),
    }
}
#[test]
fn ordinary_scalar_tags_bits_identities_and_allocation_policy() {
    let before = module(true, false);
    let after = module(false, false);
    assert_eq!(before[1].local_count, after[1].local_count);
    assert_eq!(
        before[1].autoderef_scratch_slot,
        after[1].autoderef_scratch_slot
    );
    eprintln!(
        "ordinary code {} -> {}, locals {} -> {}",
        before[1].code.len(),
        after[1].code.len(),
        before[1].local_count,
        after[1].local_count
    );
    assert!(before[1].code.len() > after[1].code.len());
    for v in [
        Value::F64(0.0),
        Value::F64(-0.0),
        Value::F64(f64::NAN),
        Value::F64(f64::INFINITY),
        Value::Bool(true),
        Value::Bool(false),
        Value::String("".into()),
        Value::String("é😀".into()),
        Value::Null,
        Value::Undefined,
    ] {
        let mut out = vec![];
        for old in [true, false] {
            let mut vm = VM::new();
            platforms::register_platforms_all(&mut vm);
            let f = vm.run(module(old, false)).unwrap();
            out.push(vm.invoke(&f, &[v.clone()]).unwrap());
            assert!(vm.last_exception.is_none());
        }
        exact(out.remove(0), out.remove(0));
    }
}
#[test]
fn ordinary_scratch_releases_prior_object_before_next_observation() {
    for old in [true, false] {
        let weak: Arc<Mutex<Option<Weak<Mutex<Object>>>>> = Arc::new(Mutex::new(None));
        let mut vm = VM::new();
        platforms::register_platforms_all(&mut vm);
        let w = weak.clone();
        vm.register_host_fn(
            "test:scalar",
            "makePrior",
            Box::new(move |_, args| {
                assert!(args.is_empty());
                let o = vybe_runtime::heap::alloc(Object::new());
                *w.lock().unwrap() = Some(Arc::downgrade(&o));
                Value::Object(o)
            }),
        );
        let w = weak.clone();
        vm.register_host_fn(
            "test:scalar",
            "observePrior",
            Box::new(move |_, args| {
                assert!(args.is_empty());
                Value::Bool(w.lock().unwrap().as_ref().unwrap().upgrade().is_some())
            }),
        );
        let f = vm.run(module(old, true)).unwrap();
        assert_eq!(
            vm.invoke(&f, &[Value::F64(3.0)]).unwrap(),
            Value::Bool(false),
            "ordinary scratch releases the prior root at the scalar write; old={old} ordinary"
        );
        // The existing top-level RETURN leaves its local vector on the VM
        // stack. Preserve that legacy observation, then require release
        // when its genuine owner is dropped.
        assert_eq!(
            weak.lock().unwrap().as_ref().unwrap().upgrade().is_some(),
            false,
            "post-return owner behavior old={old} ordinary"
        );
        drop(vm);
        assert!(
            weak.lock().unwrap().as_ref().unwrap().upgrade().is_none(),
            "no retained Object after VM owner release"
        );
    }
}
#[test]
fn ordinary_complete_native_modules_strict() {
    let mut errors = vec![];
    for old in [true, false] {
        let chunks = module(old, false);
        let wasm = vybe_platform_wasm::write_wasm(&chunks);
        let result = wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&wasm);
        if let Err(e) = result {
            errors.push(format!("old={old} ordinary: {e}"));
        }
    }
    assert!(
        errors.is_empty(),
        "unpruned complete kernel native failures: {errors:?}"
    );
}
