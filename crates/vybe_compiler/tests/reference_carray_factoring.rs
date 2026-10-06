use std::sync::{Arc, Mutex};
use vybe_compiler::primitives::{bundle, globals, platforms, pointers, references};
use vybe_runtime::value::Object;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/reference_carray_before.rs"]
mod before;

fn object(fields: &[(&str, Value)]) -> Value {
    let mut object = Object::new();
    for (name, value) in fields {
        object.properties.insert((*name).into(), value.clone());
    }
    Value::Object(vybe_runtime::heap::alloc(object))
}
fn program(old: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("<carray-factoring>");
    root.emit_op_u16(Op::REF_FUNC, 1, 0);
    root.emit(0, 0);
    root.emit_op(Op::RETURN, 0);
    let helper = if old {
        before::build_autoderef()
    } else {
        references::build_autoderef()
    };
    let mut chunks = vec![root, helper];
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
        .unwrap();
    chunks
}
fn runtime() -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm
}
fn indexed(base: Value, index: Value) -> Value {
    object(&[
        (
            pointers::REF_KIND_KEY,
            Value::String(pointers::CARRAY_KIND.into()),
        ),
        (pointers::CARRAY_BASE_KEY, base),
        (pointers::CARRAY_IDX_KEY, index),
    ])
}
#[test]
fn one_indexed_read_body_and_unchanged_local_budget() {
    let old = program(true);
    let new = program(false);
    let count = |chunk: &Chunk| {
        let (mut at, mut count) = (0, 0);
        while at < chunk.code.len() {
            let op = Op::from_code_at(&chunk.code, at);
            count += usize::from(op == Op::ARRAY_GET);
            at += 4 + op.operand_format().size_in(&chunk.code, at + 4);
        }
        count
    };
    assert_eq!(count(&old[1]), 2);
    assert_eq!(count(&new[1]), 1);
    assert_eq!(old[1].local_count, new[1].local_count);
    assert!(new[1].code.len() < old[1].code.len());
    eprintln!(
        "reference body bytes {} -> {}, locals {} -> {}",
        old[1].code.len(),
        new[1].code.len(),
        old[1].local_count,
        new[1].local_count
    );
}
#[test]
fn scalar_object_array_string_and_cell_bases_keep_results() {
    for old in [true, false] {
        let mut vm = runtime();
        let callable = vm.run(program(old)).unwrap();
        let array = Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
            Value::I32(11),
            Value::I32(23),
        ])));
        let cell = object(&[
            (
                pointers::REF_KIND_KEY,
                Value::String(pointers::CELL_KIND.into()),
            ),
            (pointers::REF_VALUE_KEY, Value::I32(37)),
        ]);
        for (input, expected) in [
            (indexed(array, Value::I32(1)), Value::I32(23)),
            (
                indexed(cell, Value::String("ignored".into())),
                Value::I32(37),
            ),
            (
                indexed(Value::String("abc".into()), Value::I32(1)),
                Value::String("b".into()),
            ),
            (Value::Null, Value::Null),
            (Value::I32(7), Value::I32(7)),
        ] {
            assert_eq!(vm.invoke(&callable, &[input]).unwrap(), expected);
        }
        let ordinary = object(&[("x", Value::I32(41))]);
        let result = vm.invoke(&callable, &[ordinary.clone()]).unwrap();
        let (Value::Object(a), Value::Object(b)) = (ordinary, result) else {
            panic!("identity")
        };
        assert!(Arc::ptr_eq(&a, &b));
    }
}
#[test]
fn fresh_getter_order_and_cell_index_skip_match_frozen() {
    for old in [true, false] {
        for cell_base in [true, false] {
            let mut vm = runtime();
            let trace = Arc::new(Mutex::new(Vec::new()));
            let getter = |vm: &mut VM, name: &'static str, value: Value| {
                let trace = trace.clone();
                vm.register_host_fn(
                    "test:carray",
                    name,
                    Box::new(move |_, _| {
                        trace.lock().unwrap().push(name);
                        value.clone()
                    }),
                );
                vm.func_table[vm.resolve_host_function_index("test:carray", name).unwrap()].clone()
            };
            let kind = getter(
                &mut vm,
                "base_kind",
                Value::String(if cell_base { "cell" } else { "other" }.into()),
            );
            let value = getter(&mut vm, "value", Value::I32(59));
            let base = object(&[
                ("__get___ref_kind", kind),
                ("__get___value", value),
                ("x", Value::I32(61)),
            ]);
            let base_get = getter(&mut vm, "base", base);
            let index_get = getter(&mut vm, "index", Value::String("x".into()));
            let input = object(&[
                (pointers::REF_KIND_KEY, Value::String("carray".into())),
                ("__get___base", base_get),
                ("__get___idx", index_get),
            ]);
            let callable = vm.run(program(old)).unwrap();
            assert_eq!(
                vm.invoke(&callable, &[input]).unwrap(),
                Value::I32(if cell_base { 59 } else { 61 })
            );
            assert_eq!(
                *trace.lock().unwrap(),
                if cell_base {
                    vec!["base", "base_kind", "value"]
                } else {
                    vec!["base", "base_kind", "index"]
                }
            );
        }
    }
}
#[test]
fn throwing_index_getter_preserves_partial_trace() {
    let mut outcomes = Vec::new();
    for old in [true, false] {
        let mut vm = runtime();
        let trace = Arc::new(Mutex::new(0));
        let observed = trace.clone();
        vm.register_host_fn(
            "test:carray",
            "throw_index",
            Box::new(move |ctx, _| {
                *observed.lock().unwrap() += 1;
                ctx.throw_value(Value::String("index-error".into()));
                Value::Null
            }),
        );
        let getter = vm.func_table[vm
            .resolve_host_function_index("test:carray", "throw_index")
            .unwrap()]
        .clone();
        let input = object(&[
            (pointers::REF_KIND_KEY, Value::String("carray".into())),
            (pointers::CARRAY_BASE_KEY, Value::String("abc".into())),
            ("__get___idx", getter),
        ]);
        let callable = vm.run(program(old)).unwrap();
        // Host throws may remain in the pending exception slot rather than
        // becoming an invoke() error. Preserve both observable channels.
        let result = vm.invoke(&callable, &[input]).map_err(|mut error| {
            for frame in &mut error.call_stack { frame.offset = 0; }
            format!("{error:?}")
        });
        outcomes.push((result, vm.last_exception.clone()));
        assert_eq!(*trace.lock().unwrap(), 1);
    }
    assert_eq!(outcomes[0], outcomes[1]);
}
