use std::sync::{Arc, Mutex};
use vybex as _;
use vybe_compiler::primitives::{bundle, callable, globals, platforms, pointers, references};
use vybe_runtime::value::Object;
use vybe_runtime::{Chunk, Op, VM, Value};

fn object(fields: &[(&str, Value)]) -> Value {
    let mut object = Object::new();
    for (key, value) in fields {
        object.properties.insert((*key).into(), value.clone());
    }
    Value::Object(vybe_runtime::heap::alloc(object))
}

fn program(shared: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("reference-read-contract");
    root.local_count = 1;
    globals::emit_read(&mut root, "input", 0);
    let mut chunks = vec![root];
    if shared {
        chunks[0].emit_op_u16(Op::LOCAL_SET, 0, 0);
        references::emit_shared_word_new(&mut chunks, 0, 0, 0);
    }
    references::emit_autoderef_to_stack(&mut chunks, 0, 0);
    chunks[0].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}

fn validate(chunks: &[Chunk]) {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(chunks))
        .expect("reference read must emit valid Wasm");
}

fn vm() -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm
}

#[test]
fn reference_reads_preserve_scalar_cell_and_array_contracts() {
    let chunks = program(false);
    validate(&chunks);
    for input in [Value::Null, Value::Bool(false), Value::I32(7), Value::String("text".into())] {
        let mut vm = vm();
        vm.set_global("input", input.clone());
        assert_eq!(vm.run(chunks.clone()).unwrap(), input);
    }
    let cell = object(&[
        (pointers::REF_KIND_KEY, Value::String(pointers::CELL_KIND.into())),
        (pointers::REF_VALUE_KEY, Value::I32(17)),
    ]);
    let array = Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
        Value::I32(3), Value::I32(29),
    ])));
    let indexed = object(&[
        (pointers::REF_KIND_KEY, Value::String(pointers::CARRAY_KIND.into())),
        (pointers::CARRAY_BASE_KEY, array),
        (pointers::CARRAY_IDX_KEY, Value::I32(1)),
    ]);
    let cell_base = object(&[
        (pointers::REF_KIND_KEY, Value::String(pointers::CARRAY_KIND.into())),
        (pointers::CARRAY_BASE_KEY, cell.clone()),
        (pointers::CARRAY_IDX_KEY, Value::I32(0)),
    ]);
    for (input, expected) in [(cell, 17), (indexed, 29), (cell_base, 17)] {
        let mut vm = vm();
        vm.set_global("input", input);
        assert_eq!(vm.run(chunks.clone()).unwrap(), Value::I32(expected));
    }
}

#[test]
fn reference_tag_getter_is_read_again_and_rechecked() {
    let chunks = program(false);
    validate(&chunks);
    for second in [Value::I32(9), Value::String("unknown".into()), Value::String("cell".into())] {
        let mut vm = vm();
        let reads = Arc::new(Mutex::new(0usize));
        let trace = reads.clone();
        let returned = second.clone();
        vm.register_host_fn("test:reference-read", "tag", Box::new(move |_, _| {
            let mut count = trace.lock().unwrap();
            *count += 1;
            if *count == 1 { Value::String("cell".into()) } else { returned.clone() }
        }));
        let getter = vm.func_table[vm.resolve_host_function_index("test:reference-read", "tag").unwrap()].clone();
        let input = object(&[
            ("__get___ref_kind", getter),
            (pointers::REF_VALUE_KEY, Value::I32(41)),
        ]);
        vm.set_global("input", input.clone());
        let result = vm.run(chunks.clone()).unwrap();
        assert_eq!(*reads.lock().unwrap(), 2, "do not cache observable tag reads");
        if second == Value::String("cell".into()) {
            assert_eq!(result, Value::I32(41));
        } else {
            let (Value::Object(before), Value::Object(after)) = (&input, &result) else {
                panic!("non-reference tag must preserve object identity");
            };
            assert!(Arc::ptr_eq(before, after));
        }
    }
}

#[test]
fn shared_reference_reads_use_the_shared_word() {
    let chunks = program(true);
    validate(&chunks);
    let mut vm = vm();
    vm.set_global("input", Value::I32(53));
    assert_eq!(vm.run(chunks).unwrap(), Value::I32(53));
}

#[test]
fn guarded_primitive_body_keeps_tag_checks_and_is_smaller() {
    let guarded = references::build_guarded_reference_read();
    let original = references::build_autoderef();
    assert!(guarded.code.len() < original.code.len());
    eprintln!("reference-read body bytes: {} -> {}", original.code.len(), guarded.code.len());

    let mut root = Chunk::new("guarded-reference-read-contract");
    root.local_count = 1;
    globals::emit_read(&mut root, "input", 0);
    root.emit_op_u16(Op::LOCAL_SET, 0, 0);
    root.emit_op_u16(Op::REF_FUNC, 1, 0);
    root.emit(0, 0);
    root.emit_op_u16(Op::LOCAL_GET, 0, 0);
    callable::emit_direct_invoke_chunk(&mut root, 1, 0);
    root.emit_op(Op::RETURN, 0);
    let mut chunks = vec![root, guarded];
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    validate(&chunks);

    for tag in [Value::I32(9), Value::String("unknown".into()), Value::String("cell".into())] {
        let mut vm = vm();
        let reads = Arc::new(Mutex::new(0usize));
        let trace = reads.clone();
        let returned = tag.clone();
        vm.register_host_fn("test:guarded-reference-read", "tag", Box::new(move |_, _| {
            *trace.lock().unwrap() += 1;
            returned.clone()
        }));
        let getter = vm.func_table[vm.resolve_host_function_index("test:guarded-reference-read", "tag").unwrap()].clone();
        let input = object(&[
            ("__get___ref_kind", getter),
            (pointers::REF_VALUE_KEY, Value::I32(41)),
        ]);
        vm.set_global("input", input.clone());
        let result = vm.run(chunks.clone()).unwrap();
        assert_eq!(*reads.lock().unwrap(), 1, "body must perform its own fresh tag read");
        if tag == Value::String("cell".into()) {
            assert_eq!(result, Value::I32(41));
        } else {
            let (Value::Object(before), Value::Object(after)) = (&input, &result) else {
                panic!("non-reference tag must preserve object identity");
            };
            assert!(Arc::ptr_eq(before, after));
        }
    }
}

#[test]
fn guarded_inline_read_does_not_retain_getter_produced_base() {
    let mut vm = vm();
    let base = object(&[
        (pointers::REF_KIND_KEY, Value::String(pointers::CELL_KIND.into())),
        (pointers::REF_VALUE_KEY, Value::I32(37)),
    ]);
    let returned = base.clone();
    vm.register_host_fn("test:reference-lifetime", "base", Box::new(move |_, _| returned.clone()));
    let observed = base.clone();
    vm.register_host_fn("test:reference-lifetime", "reachable", Box::new(move |ctx, _| {
        Value::Bool(ctx.is_reachable_from_program(&observed))
    }));
    let getter = vm.func_table[vm.resolve_host_function_index("test:reference-lifetime", "base").unwrap()].clone();
    vm.set_global("input", object(&[
        (pointers::REF_KIND_KEY, Value::String(pointers::CARRAY_KIND.into())),
        ("__get___base", getter),
        (pointers::CARRAY_IDX_KEY, Value::I32(0)),
    ]));

    let mut root = Chunk::new("guarded-reference-inline-lifetime");
    root.local_count = 1;
    globals::emit_read(&mut root, "input", 0);
    root.emit_op_u16(Op::LOCAL_SET, 0, 0);
    let mut chunks = vec![root];
    references::emit_guarded_reference_read_inline_from_slot(&mut chunks, 0, 0, 0);
    chunks[0].emit_op(Op::DROP, 0);
    let reachable = chunks[0].add_import("test:reference-lifetime", "reachable");
    chunks[0].emit_call(reachable, 0, 0);
    chunks[0].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    validate(&chunks);
    assert_eq!(vm.run(chunks).unwrap(), Value::Bool(false));
}
