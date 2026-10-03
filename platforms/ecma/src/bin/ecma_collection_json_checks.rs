//! Focused ECMA adapter checks through an actual Ambient VM call frame.
use std::collections::HashMap;
use vybe_platform_ecma as ecma;
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Chunk, HostContext, Op, VM, Value};

fn host_ref(vm: &VM, module: &str, name: &str) -> Value {
    let mut object = Object::new();
    object.kind = ObjectKind::HostFunction(vm.resolve_host_function_index(module, name).unwrap());
    Value::Object(vybe_runtime::heap::alloc(object))
}

fn array(values: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)))
}

fn string(text: &str) -> Value {
    ecma::keys::string_value(text)
}

fn assert_text(value: Value, expected: &str) {
    let Value::String(actual) = value else {
        panic!("expected string")
    };
    assert_eq!(actual.as_ref(), expected);
}

fn field(value: &Value, key: &str) -> Value {
    let Value::Object(object) = value else {
        panic!("expected object")
    };
    object
        .lock()
        .unwrap()
        .properties
        .get(key)
        .cloned()
        .unwrap_or(Value::Undefined)
}

fn numbers(value: &Value) -> Vec<i32> {
    let Value::Object(object) = value else {
        panic!("expected array")
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(values) = &object.kind else {
        panic!("expected array backing")
    };
    values.iter().map(Value::as_i32).collect()
}

fn invoke(
    ctx: &mut HostContext,
    imports: &HashMap<&str, Value>,
    name: &str,
    args: &[Value],
) -> Value {
    let result = ctx.try_invoke(&imports[name], args)
        .unwrap_or_else(|error| panic!("{name} threw: {error:?}"));
    // These typed imports declare i32 predicates at the native WASM boundary.
    // Check their encoding before expressing the assertion as a Rust boolean.
    if matches!(name, "isSubsetOf" | "isSupersetOf" | "isDisjointFrom") {
        match result {
            Value::I32(0) => Value::Bool(false),
            Value::I32(1) => Value::Bool(true),
            other => panic!("{name} returned an invalid i32 predicate: {other:?}"),
        }
    } else {
        result
    }
}

fn main() {
    let mut vm = VM::new();
    ecma::register(&mut vm);
    vm.register_free_fn(
        "collection-check",
        "double",
        Box::new(|_, args| Value::I32(args[0].as_i32() * 2)),
    );
    vm.register_host_fn(
        "collection-check",
        "holder",
        Box::new(|_, args| {
            let Value::Object(holder) = &args[0] else {
                panic!("missing replacer holder")
            };
            // Fails immediately rather than hanging if stringify retains this lock.
            let guard = holder.try_lock().expect("replacer holder must be unlocked");
            drop(guard);
            args[2].clone()
        }),
    );
    vm.register_host_fn(
        "collection-check",
        "revive",
        Box::new(|_, args| match &args[2] {
            Value::I32(number) => Value::I32(number * 2),
            other => other.clone(),
        }),
    );
    let mapper = host_ref(&vm, "collection-check", "double");
    let holder = ecma::receiver_host_fn_ref(
        "collection-check",
        "holder",
        vm.resolve_host_function_index("collection-check", "holder")
            .unwrap(),
    );
    let reviver = ecma::receiver_host_fn_ref(
        "collection-check",
        "revive",
        vm.resolve_host_function_index("collection-check", "revive")
            .unwrap(),
    );
    let mut imports = HashMap::new();
    for (module, names) in [
        (
            "ecma:set",
            &[
                "isSubsetOf",
                "isSupersetOf",
                "isDisjointFrom",
                "overlaps",
                "difference",
            ][..],
        ),
        ("ecma:iterator", &["from", "next", "map", "toArray"][..]),
        ("ecma:string", &["slice", "replace", "includes"][..]),
        ("ecma:json", &["parse", "stringify"][..]),
        ("ecma:arraybuffer", &["newWithLength", "newResizable", "byteLength", "maxByteLength", "resizable", "detached", "transfer"][..]),
    ] {
        for name in names {
            imports.insert(*name, host_ref(&vm, module, name));
        }
    }
    vm.register_free_fn(
        "collection-check",
        "run",
        Box::new(move |ctx, _| {
            let set = |values: Vec<i32>| {
                ecma::set::make_set(values.into_iter().map(Value::I32).collect())
            };
            let empty = set(vec![]);
            let small = set(vec![2, 3]);
            let large = set((0..2048).collect());
            let separate = set(vec![4096, 8192]);
            for (left, right) in [(&small, &large), (&large, &small)] {
                assert!(
                    !invoke(
                        ctx,
                        &imports,
                        "isDisjointFrom",
                        &[left.clone(), right.clone()]
                    )
                    .as_bool()
                );
                assert!(
                    invoke(ctx, &imports, "overlaps", &[left.clone(), right.clone()]).as_bool()
                );
            }
            assert!(invoke(ctx, &imports, "isSubsetOf", &[small.clone(), large.clone()]).as_bool());
            assert!(
                !invoke(ctx, &imports, "isSubsetOf", &[large.clone(), small.clone()]).as_bool()
            );
            assert!(
                invoke(
                    ctx,
                    &imports,
                    "isSupersetOf",
                    &[large.clone(), small.clone()]
                )
                .as_bool()
            );
            assert!(
                !invoke(
                    ctx,
                    &imports,
                    "isSupersetOf",
                    &[small.clone(), large.clone()]
                )
                .as_bool()
            );
            for (left, right) in [(&large, &separate), (&separate, &large), (&empty, &large)] {
                assert!(
                    invoke(
                        ctx,
                        &imports,
                        "isDisjointFrom",
                        &[left.clone(), right.clone()]
                    )
                    .as_bool()
                );
                assert!(
                    !invoke(ctx, &imports, "overlaps", &[left.clone(), right.clone()]).as_bool()
                );
            }
            assert!(invoke(ctx, &imports, "isSubsetOf", &[small.clone(), small.clone()]).as_bool());
            let difference = invoke(ctx, &imports, "difference", &[large.clone(), small.clone()]);
            let Value::Object(difference) = difference else {
                panic!("missing difference set")
            };
            let guard = difference.lock().unwrap();
            let ObjectKind::Set(values) = &guard.kind else {
                panic!("missing set backing")
            };
            assert_eq!(values.len(), 2046);
            assert_eq!(
                values.iter().take(4).map(Value::as_i32).collect::<Vec<_>>(),
                vec![0, 1, 4, 5]
            );
            drop(guard);

            let source = array(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
            let iterator = invoke(ctx, &imports, "from", &[source.clone()]);
            let first = invoke(ctx, &imports, "next", &[iterator.clone()]);
            assert_eq!(field(&first, "value").as_i32(), 1);
            assert!(!field(&first, "done").as_bool());
            assert_eq!(
                numbers(&invoke(ctx, &imports, "toArray", &[iterator])),
                vec![2, 3]
            );
            let mapped = invoke(ctx, &imports, "map", &[source, mapper.clone()]);
            let first = invoke(ctx, &imports, "next", &[mapped.clone()]);
            assert_eq!(field(&first, "value").as_i32(), 2);
            assert_eq!(
                numbers(&invoke(ctx, &imports, "toArray", &[mapped.clone()])),
                vec![4, 6]
            );
            assert!(field(&invoke(ctx, &imports, "next", &[mapped]), "done").as_bool());

            let parsed = invoke(
                ctx,
                &imports,
                "parse",
                &[string(r#"{"a":[1,true,null],"b":"x"}"#)],
            );
            let output = invoke(ctx, &imports, "stringify", &[parsed]);
            assert_text(output, r#"{"a":[1,true,null],"b":"x"}"#);
            let mut escaped = Object::new();
            escaped
                .properties
                .insert("quote\"\\\n".into(), string("line\tend"));
            let escaped = Value::Object(vybe_runtime::heap::alloc(escaped));
            assert_text(
                invoke(ctx, &imports, "stringify", &[escaped]),
                r#"{"quote\"\\\n":"line\tend"}"#
            );
            assert_text(
                invoke(ctx, &imports, "stringify", &[array(vec![])]),
                "[]"
            );
            let values = array(vec![Value::I32(1), Value::I32(2)]);
            assert_text(
                invoke(ctx, &imports, "stringify", &[values, holder.clone()]),
                "[1,2]"
            );
            let parsed = invoke(
                ctx,
                &imports,
                "parse",
                &[string(r#"{"a":1,"b":[2,3]}"#), reviver.clone()],
            );
            assert_eq!(field(&parsed, "a").as_i32(), 2);
            assert_eq!(numbers(&field(&parsed, "b")), vec![4, 6]);
            let Value::Object(cycle) = array(vec![]) else {
                unreachable!()
            };
            {
                let mut guard = cycle.lock().unwrap();
                let ObjectKind::Array(values) = &mut guard.kind else {
                    unreachable!()
                };
                values.push(Value::Object(cycle.clone()));
            }
            let error = ctx.try_invoke(&imports["stringify"], &[Value::Object(cycle.clone())]);
            {
                let mut guard = cycle.lock().unwrap();
                let ObjectKind::Array(values) = &mut guard.kind else {
                    unreachable!()
                };
                values.clear();
            }
            assert!(error.is_err(), "recursive arrays must throw");
            let buffer = invoke(ctx, &imports, "newWithLength", &[Value::I32(8)]);
            assert_eq!(invoke(ctx, &imports, "byteLength", &[buffer.clone()]).as_i32(), 8);
            assert_eq!(invoke(ctx, &imports, "maxByteLength", &[buffer.clone()]).as_i32(), 8);
            assert!(!invoke(ctx, &imports, "resizable", &[buffer.clone()]).as_bool());
            assert!(!invoke(ctx, &imports, "detached", &[buffer.clone()]).as_bool());
            let transferred = invoke(ctx, &imports, "transfer", &[buffer.clone()]);
            assert_eq!(invoke(ctx, &imports, "byteLength", &[buffer.clone()]).as_i32(), 0);
            assert!(invoke(ctx, &imports, "detached", &[buffer]).as_bool());
            assert_eq!(invoke(ctx, &imports, "byteLength", &[transferred]).as_i32(), 8);
            let resizable = invoke(ctx, &imports, "newResizable", &[Value::I32(4), Value::I32(16)]);
            assert_eq!(invoke(ctx, &imports, "byteLength", &[resizable.clone()]).as_i32(), 4);
            assert_eq!(invoke(ctx, &imports, "maxByteLength", &[resizable.clone()]).as_i32(), 16);
            assert!(invoke(ctx, &imports, "resizable", &[resizable]).as_bool());
            for (source, needle, position, expected) in [
                ("A\u{1f600}B", "\u{1f600}", 0, true),
                ("A\u{1f600}B", "\u{1f600}", -3, true),
                ("A\u{1f600}B", "\u{1f600}", 1, true),
                ("A\u{1f600}B", "\u{1f600}", 2, false),
                ("A\u{1f600}B", "B", 2, true),
                ("A\u{1f600}B", "B", 4, false),
                ("caf\u{e9}", "", 99, true),
                ("caf\u{e9}", "z", 0, false),
            ] {
                assert_eq!(
                    invoke(ctx, &imports, "includes", &[string(source), string(needle), Value::I32(position)]).as_bool(),
                    expected,
                );
            }
            for (source, start, end, expected) in [
                ("abcdef", 1, 4, "bcd"),
                ("abcdef", -3, -1, "de"),
                ("abcdef", -99, 99, "abcdef"),
                ("abcdef", 4, 2, ""),
                ("", 0, 1, ""),
                ("A\u{1f600}B", 1, 3, "\u{1f600}"),
            ] {
                assert_text(
                    invoke(ctx, &imports, "slice", &[string(source), Value::I32(start), Value::I32(end)]),
                    expected,
                );
            }
            assert_text(invoke(ctx, &imports, "slice", &[string("abcdef"), Value::I32(-2)]), "ef");
            for (source, search, replacement, expected) in [
                ("abab", "ab", "x", "xab"),
                ("ab", "", "-", "-ab"),
                ("ab", "z", "x", "ab"),
                ("caf\u{e9}", "\u{e9}", "e", "cafe"),
            ] {
                assert_text(
                    invoke(ctx, &imports, "replace", &[string(source), string(search), string(replacement)]),
                    expected,
                );
            }
            Value::Bool(true)
        }),
    );
    let mut entry = Chunk::new("collection-check");
    entry.local_count = 1;
    let run = entry.add_import("collection-check", "run");
    entry.emit_call(run, 0, 1);
    entry.emit_op(Op::RETURN, 1);
    assert!(matches!(
        vm.run(vec![entry]).expect("checks must complete"),
        Value::Bool(true)
    ));
    println!(
        "Set predicates/order, iterator cursors, JSON compact/escaping/reviver/holder/cycle checks passed"
    );
}
