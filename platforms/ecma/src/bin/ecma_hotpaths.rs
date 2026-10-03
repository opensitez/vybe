//! Direct ECMA host-surface benchmark. Registration and fixture creation are
//! excluded; VM host dispatch is included. No source compilation is performed.
use std::hint::black_box;
use std::sync::{
    Arc,
    atomic::{AtomicI32, Ordering},
};
use std::time::Instant;
use vybe_platform_ecma as ecma;
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{VM, Value};

fn host_ref(vm: &VM, module: &str, name: &str) -> Value {
    let index = vm
        .resolve_host_function_index(module, name)
        .unwrap_or_else(|| panic!("missing builtin {module}.{name}"));
    let mut object = Object::new();
    object.kind = ObjectKind::HostFunction(index);
    Value::Object(vybe_runtime::heap::alloc(object))
}

fn method_ref(vm: &VM, name: &str) -> Value {
    let index = vm.resolve_host_function_index("benchmark", name).unwrap();
    ecma::receiver_host_fn_ref("benchmark", name, index)
}

fn object_with(properties: &[(&str, Value)]) -> Value {
    let mut object = Object::new();
    object.properties.reserve(properties.len());
    for (key, value) in properties {
        object.properties.insert((*key).into(), value.clone());
    }
    Value::Object(vybe_runtime::heap::alloc(object))
}

fn invoke(vm: &mut VM, target: &Value, args: &[Value]) -> Value {
    vm.invoke(target, args)
        .expect("benchmark invocation failed")
}

fn numeric_case(
    vm: &mut VM,
    label: &str,
    target: &Value,
    args: &[Value],
    expected: f64,
    iterations: usize,
) {
    assert_eq!(invoke(vm, target, args).as_f64(), expected, "{label}");
    for _ in 0..iterations.min(1_000) {
        black_box(invoke(vm, target, black_box(args)));
    }
    let start = Instant::now();
    for _ in 0..iterations {
        black_box(invoke(vm, target, black_box(args)));
    }
    let elapsed = start.elapsed();
    println!(
        "{label:24} {:10.1} ns/call  {:9.3} ms  n={iterations}",
        elapsed.as_secs_f64() * 1e9 / iterations as f64,
        elapsed.as_secs_f64() * 1e3,
    );
}

fn boolean_case(
    vm: &mut VM,
    label: &str,
    target: &Value,
    args: &[Value],
    expected: bool,
    iterations: usize,
) {
    assert!(
        matches!(invoke(vm, target, args), Value::Bool(value) if value == expected),
        "{label}"
    );
    for _ in 0..iterations.min(1_000) {
        black_box(invoke(vm, target, black_box(args)));
    }
    let start = Instant::now();
    for _ in 0..iterations {
        black_box(invoke(vm, target, black_box(args)));
    }
    let elapsed = start.elapsed();
    println!(
        "{label:24} {:10.1} ns/call  {:9.3} ms  n={iterations}",
        elapsed.as_secs_f64() * 1e9 / iterations as f64,
        elapsed.as_secs_f64() * 1e3,
    );
}

fn main() {
    let iterations = std::env::args()
        .nth(1)
        .map(|arg| arg.parse::<usize>().expect("iterations must be an integer"))
        .unwrap_or(100_000);
    assert!(iterations > 0, "iterations must be positive");
    let mut vm = VM::new();
    ecma::register(&mut vm);
    vm.register_host_fn(
        "benchmark",
        "sum",
        Box::new(|_, args| {
            Value::F64(args.get(1..).unwrap_or(&[]).iter().map(Value::as_f64).sum())
        }),
    );
    vm.register_free_fn("benchmark", "getter", Box::new(|_, _| Value::I32(73)));
    let setter_result = Arc::new(AtomicI32::new(0));
    let setter_capture = setter_result.clone();
    vm.register_host_fn(
        "benchmark",
        "setter",
        Box::new(move |_, args| {
            setter_capture.store(args.last().unwrap().as_f64() as i32, Ordering::Relaxed);
            Value::Undefined
        }),
    );

    let get = host_ref(&vm, "ecma:object", "get");
    let has = host_ref(&vm, "ecma:object", "has");
    let set = host_ref(&vm, "ecma:object", "set");
    let call = host_ref(&vm, "ecma:function", "call");
    let apply = host_ref(&vm, "ecma:function", "apply");
    let reflect_apply = host_ref(&vm, "ecma:reflect", "apply");
    let sum = method_ref(&vm, "sum");
    let key = ecma::keys::string_value("value");
    let own = object_with(&[("__proto__", Value::Null), ("value", Value::I32(42))]);
    let child = object_with(&[("__proto__", own.clone())]);
    boolean_case(
        &mut vm,
        "object.has.own",
        &has,
        &[own.clone(), key.clone()],
        true,
        iterations,
    );
    boolean_case(
        &mut vm,
        "object.has.inherited",
        &has,
        &[child.clone(), key.clone()],
        true,
        iterations,
    );
    boolean_case(
        &mut vm,
        "object.has.missing",
        &has,
        &[child.clone(), ecma::keys::string_value("absent")],
        false,
        iterations,
    );
    boolean_case(
        &mut vm,
        "object.has.undefined",
        &has,
        &[
            object_with(&[("__proto__", Value::Null), ("value", Value::Undefined)]),
            key.clone(),
        ],
        true,
        iterations,
    );
    boolean_case(
        &mut vm,
        "object.has.numeric",
        &has,
        &[
            object_with(&[("__proto__", Value::Null), ("7", Value::I32(1))]),
            Value::I32(7),
        ],
        true,
        iterations,
    );
    let getter = object_with(&[
        ("__proto__", Value::Null),
        ("__get_value", host_ref(&vm, "benchmark", "getter")),
    ]);
    assert_eq!(
        invoke(&mut vm, &get, &[getter.clone(), key.clone()]).as_f64(),
        73.0
    );
    let getter_child = object_with(&[("__proto__", getter)]);
    invoke(
        &mut vm,
        &set,
        &[getter_child.clone(), key.clone(), Value::I32(99)],
    );
    assert_eq!(
        invoke(&mut vm, &get, &[getter_child.clone(), key.clone()]).as_f64(),
        73.0
    );
    if let Value::Object(object) = getter_child {
        assert!(!object.lock().unwrap().properties.contains_key("value"));
    }
    let setter_prototype = object_with(&[
        ("__proto__", Value::Null),
        ("__set_value", host_ref(&vm, "benchmark", "setter")),
    ]);
    let setter_child = object_with(&[("__proto__", setter_prototype)]);
    invoke(
        &mut vm,
        &set,
        &[setter_child.clone(), key.clone(), Value::I32(99)],
    );
    assert_eq!(setter_result.load(Ordering::Relaxed), 99);
    if let Value::Object(object) = setter_child {
        assert!(!object.lock().unwrap().properties.contains_key("value"));
    }
    let shadow = object_with(&[("__proto__", child.clone()), ("value", Value::Undefined)]);
    assert!(matches!(
        invoke(&mut vm, &get, &[shadow, key.clone()]),
        Value::Undefined
    ));

    println!(
        "ECMA host calls; setup excluded, VM dispatch included; profile={}",
        if cfg!(debug_assertions) {
            "debug"
        } else {
            "optimized"
        }
    );
    numeric_case(
        &mut vm,
        "object.get.own",
        &get,
        &[own.clone(), key.clone()],
        42.0,
        iterations,
    );
    numeric_case(
        &mut vm,
        "object.get.inherited",
        &get,
        &[child, key.clone()],
        42.0,
        iterations,
    );

    let write_args = [own.clone(), key.clone(), Value::I32(42)];
    invoke(&mut vm, &set, &write_args);
    assert_eq!(invoke(&mut vm, &get, &[own, key]).as_f64(), 42.0);
    for _ in 0..iterations.min(1_000) {
        black_box(invoke(&mut vm, &set, &write_args));
    }
    let start = Instant::now();
    for _ in 0..iterations {
        black_box(invoke(&mut vm, &set, black_box(&write_args)));
    }
    let elapsed = start.elapsed();
    println!(
        "{:<24} {:10.1} ns/call  {:9.3} ms  n={iterations}",
        "object.set.existing",
        elapsed.as_secs_f64() * 1e9 / iterations as f64,
        elapsed.as_secs_f64() * 1e3,
    );

    let array = Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
        Value::I32(
            42
        );
        32
    ])));
    numeric_case(
        &mut vm,
        "object.get.array-index",
        &get,
        &[array, Value::I32(16)],
        42.0,
        iterations,
    );
    numeric_case(
        &mut vm,
        "function.call.4",
        &call,
        &[
            sum.clone(),
            Value::Null,
            Value::I32(1),
            Value::I32(2),
            Value::I32(3),
            Value::I32(4),
        ],
        10.0,
        iterations,
    );
    for arity in [0usize, 4, 8, 9, 16] {
        let arguments = Value::Object(vybe_runtime::heap::alloc(Object::new_array(
            (1..=arity).map(|n| Value::I32(n as i32)).collect(),
        )));
        numeric_case(
            &mut vm,
            &format!("function.apply.{arity}"),
            &apply,
            &[sum.clone(), Value::Null, arguments.clone()],
            (arity * (arity + 1) / 2) as f64,
            iterations,
        );
        numeric_case(
            &mut vm,
            &format!("reflect.apply.{arity}"),
            &reflect_apply,
            &[sum.clone(), Value::Null, arguments],
            (arity * (arity + 1) / 2) as f64,
            iterations,
        );
    }
    vm.register_host_fn(
        "benchmark",
        "weighted",
        Box::new(|_, args| {
            Value::F64(
                args.get(1..)
                    .unwrap_or(&[])
                    .iter()
                    .enumerate()
                    .map(|(i, value)| (i + 1) as f64 * value.as_f64())
                    .sum(),
            )
        }),
    );
    let weighted = method_ref(&vm, "weighted");
    for arity in [0usize, 8, 9, 16] {
        let mut array_like = Object::new();
        array_like
            .properties
            .insert("length".into(), Value::I32(arity as i32));
        for index in 0..arity {
            array_like
                .properties
                .insert(index.to_string(), Value::I32(index as i32 + 1));
        }
        let array_like = Value::Object(vybe_runtime::heap::alloc(array_like));
        let expected = (arity * (arity + 1) * (2 * arity + 1) / 6) as f64;
        for method in [&apply, &reflect_apply] {
            assert_eq!(
                invoke(
                    &mut vm,
                    method,
                    &[weighted.clone(), Value::Null, array_like.clone()]
                )
                .as_f64(),
                expected
            );
        }
        let dense = Value::Object(vybe_runtime::heap::alloc(Object::new_array(
            (1..=arity).map(|n| Value::I32(n as i32)).collect(),
        )));
        assert_eq!(
            invoke(&mut vm, &apply, &[weighted.clone(), Value::Null, dense]).as_f64(),
            expected
        );
    }
    let mutation_source = vybe_runtime::heap::alloc(Object::new_array(vec![Value::I32(1); 16]));
    let mutation_capture = mutation_source.clone();
    vm.register_host_fn(
        "benchmark",
        "mutateSource",
        Box::new(move |_, args| {
            let mut source = mutation_capture.lock().unwrap();
            if let ObjectKind::Array(values) = &mut source.kind {
                values.clear();
            }
            Value::I32(args.len().saturating_sub(1) as i32)
        }),
    );
    let mutate_source = method_ref(&vm, "mutateSource");
    assert_eq!(
        invoke(
            &mut vm,
            &apply,
            &[mutate_source, Value::Null, Value::Object(mutation_source)]
        )
        .as_f64(),
        16.0
    );
    let trace = Arc::new(std::sync::Mutex::new(Vec::<String>::new()));
    let observed = vybe_runtime::heap::alloc(Object::new());
    let length_trace = trace.clone();
    vm.register_host_fn(
        "benchmark",
        "listLength",
        Box::new(move |_, _| {
            length_trace.lock().unwrap().push("length".into());
            Value::I32(3)
        }),
    );
    let first_trace = trace.clone();
    let first_source = observed.clone();
    vm.register_host_fn(
        "benchmark",
        "listFirst",
        Box::new(move |_, _| {
            first_trace.lock().unwrap().push("0".into());
            first_source
                .lock()
                .unwrap()
                .properties
                .insert("1".into(), Value::I32(20));
            Value::I32(1)
        }),
    );
    let last_trace = trace.clone();
    vm.register_host_fn(
        "benchmark",
        "listLast",
        Box::new(move |_, _| {
            last_trace.lock().unwrap().push("2".into());
            Value::I32(3)
        }),
    );
    {
        let mut source = observed.lock().unwrap();
        source.properties.insert("__proto__".into(), Value::Null);
        source
            .properties
            .insert("__get_length".into(), method_ref(&vm, "listLength"));
        source
            .properties
            .insert("__get_0".into(), method_ref(&vm, "listFirst"));
        source
            .properties
            .insert("__get_2".into(), method_ref(&vm, "listLast"));
    }
    let inherited = object_with(&[("__proto__", Value::Null), ("0", Value::I32(1))]);
    let inherited_list = object_with(&[
        ("__proto__", inherited),
        ("length", Value::I32(3)),
        ("1", Value::I32(2)),
        ("2", Value::I32(3)),
    ]);
    let mut sparse = Object::new_array(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
    sparse.properties.insert(
        "__proto__".into(),
        object_with(&[("__proto__", Value::Null), ("1", Value::I32(20))]),
    );
    ecma::array::mark_array_hole(&mut sparse, 1);
    let sparse = Value::Object(vybe_runtime::heap::alloc(sparse));
    for method in [&apply, &reflect_apply] {
        trace.lock().unwrap().clear();
        observed
            .lock()
            .unwrap()
            .properties
            .insert("1".into(), Value::I32(2));
        assert_eq!(
            invoke(
                &mut vm,
                method,
                &[
                    weighted.clone(),
                    Value::Null,
                    Value::Object(observed.clone())
                ]
            )
            .as_f64(),
            50.0
        );
        assert_eq!(*trace.lock().unwrap(), vec!["length", "0", "2"]);
        assert_eq!(
            invoke(
                &mut vm,
                method,
                &[weighted.clone(), Value::Null, inherited_list.clone()]
            )
            .as_f64(),
            14.0
        );
        assert_eq!(
            invoke(
                &mut vm,
                method,
                &[weighted.clone(), Value::Null, sparse.clone()]
            )
            .as_f64(),
            50.0
        );
    }
    let proxy_trace = trace.clone();
    vm.register_host_fn(
        "benchmark",
        "listProxyGet",
        Box::new(move |_, args| {
            let Value::String(key) = &args[2] else {
                panic!("proxy key must be a string")
            };
            proxy_trace.lock().unwrap().push(key.to_string());
            match key.as_ref() {
                "length" => Value::I32(3),
                "0" => Value::I32(1),
                "1" => Value::I32(2),
                "2" => Value::I32(3),
                _ => Value::Undefined,
            }
        }),
    );
    let proxy_new = host_ref(&vm, "ecma:proxy", "new");
    let proxy_handler = object_with(&[("get", method_ref(&vm, "listProxyGet"))]);
    let proxy_source = invoke(
        &mut vm,
        &proxy_new,
        &[object_with(&[("__proto__", Value::Null)]), proxy_handler],
    );
    for method in [&apply, &reflect_apply] {
        trace.lock().unwrap().clear();
        assert_eq!(
            invoke(
                &mut vm,
                method,
                &[weighted.clone(), Value::Null, proxy_source.clone()]
            )
            .as_f64(),
            14.0
        );
        assert_eq!(*trace.lock().unwrap(), vec!["length", "0", "1", "2"]);
    }
    let calls = Arc::new(AtomicI32::new(0));
    let calls_capture = calls.clone();
    vm.register_host_fn(
        "benchmark",
        "mustNotRun",
        Box::new(move |_, _| {
            calls_capture.fetch_add(1, Ordering::Relaxed);
            Value::Undefined
        }),
    );
    vm.register_host_fn(
        "benchmark",
        "throwElement",
        Box::new(|ctx, _| {
            ctx.throw_value(Value::I32(123));
            Value::Undefined
        }),
    );
    vm.register_free_fn(
        "benchmark",
        "catchApply",
        Box::new(|ctx, args| {
            let method = &args[0];
            match ctx.try_invoke(method, &args[1..]) {
                Err(value) => value,
                Ok(_) => Value::I32(-1),
            }
        }),
    );
    let throwing_list = object_with(&[
        ("__proto__", Value::Null),
        ("length", Value::I32(1)),
        ("__get_0", method_ref(&vm, "throwElement")),
    ]);
    let catch_apply = host_ref(&vm, "benchmark", "catchApply");
    let must_not_run = method_ref(&vm, "mustNotRun");
    for method in [&apply, &reflect_apply] {
        assert_eq!(
            invoke(
                &mut vm,
                &catch_apply,
                &[
                    method.clone(),
                    must_not_run.clone(),
                    Value::Null,
                    throwing_list.clone()
                ]
            )
            .as_f64(),
            123.0
        );
    }
    assert_eq!(calls.load(Ordering::Relaxed), 0);
    vm.register_host_fn(
        "benchmark",
        "countArgs",
        Box::new(|_, args| Value::I32(args.len().saturating_sub(1) as i32)),
    );
    let count_args = method_ref(&vm, "countArgs");
    for (length, expected) in [
        (Value::F64(2.9), 2.0),
        (Value::F64(-2.9), 0.0),
        (Value::F64(f64::NAN), 0.0),
        (Value::Bool(true), 1.0),
        (Value::Bool(false), 0.0),
        (Value::Null, 0.0),
        (Value::Undefined, 0.0),
        (ecma::keys::string_value(" 2.9 "), 2.0),
        (ecma::keys::string_value("2e0"), 2.0),
        (ecma::keys::string_value("0x2"), 2.0),
        (ecma::keys::string_value("0b10"), 2.0),
        (ecma::keys::string_value("not a number"), 0.0),
    ] {
        let list = object_with(&[
            ("length", length),
            ("0", Value::I32(1)),
            ("1", Value::I32(2)),
        ]);
        for method in [&apply, &reflect_apply] {
            assert_eq!(
                invoke(
                    &mut vm,
                    method,
                    &[count_args.clone(), Value::Null, list.clone()]
                )
                .as_f64(),
                expected
            );
        }
    }
    vm.register_host_fn("benchmark", "lengthValueOf", Box::new(|_, _| Value::I32(2)));
    vm.register_host_fn(
        "benchmark",
        "lengthPrimitive",
        Box::new(|_, args| {
            assert!(matches!(&args[1], Value::String(hint) if hint.as_ref() == "number"));
            Value::I32(2)
        }),
    );
    vm.register_host_fn("benchmark", "nullValueOf", Box::new(|_, _| Value::Null));
    for (length, expected) in [
        (
            object_with(&[("valueOf", method_ref(&vm, "lengthValueOf"))]),
            2.0,
        ),
        (
            object_with(&[("toprimitive", method_ref(&vm, "lengthPrimitive"))]),
            2.0,
        ),
        (
            object_with(&[("valueOf", method_ref(&vm, "nullValueOf"))]),
            0.0,
        ),
    ] {
        let list = object_with(&[
            ("length", length),
            ("0", Value::I32(1)),
            ("1", Value::I32(2)),
        ]);
        for method in [&apply, &reflect_apply] {
            assert_eq!(
                invoke(
                    &mut vm,
                    method,
                    &[count_args.clone(), Value::Null, list.clone()]
                )
                .as_f64(),
                expected
            );
        }
    }
    let symbol_new = host_ref(&vm, "ecma:symbol", "new");
    let symbol_length = invoke(&mut vm, &symbol_new, &[]);
    let bigint_length = Value::BigInt(ecma::bigint::parse_bigint_str("1").unwrap().into());
    for length in [symbol_length, bigint_length] {
        let list = object_with(&[("length", length)]);
        for method in [&apply, &reflect_apply] {
            assert_apply_error(
                &mut vm,
                &catch_apply,
                &get,
                method,
                &must_not_run,
                list.clone(),
                "TypeError",
            );
        }
    }
    for list in [
        Value::I32(1),
        Value::Bool(true),
        ecma::keys::string_value("abc"),
    ] {
        for method in [&apply, &reflect_apply] {
            assert_apply_error(
                &mut vm,
                &catch_apply,
                &get,
                method,
                &must_not_run,
                list.clone(),
                "TypeError",
            );
        }
    }
    for list in [Value::Null, Value::Undefined] {
        assert_eq!(
            invoke(
                &mut vm,
                &apply,
                &[count_args.clone(), Value::Null, list.clone()]
            )
            .as_f64(),
            0.0
        );
        assert_apply_error(
            &mut vm,
            &catch_apply,
            &get,
            &reflect_apply,
            &must_not_run,
            list,
            "TypeError",
        );
    }
    let throwing_length = object_with(&[("valueOf", method_ref(&vm, "throwElement"))]);
    let throwing_length_list = object_with(&[("length", throwing_length)]);
    for method in [&apply, &reflect_apply] {
        assert_eq!(
            invoke(
                &mut vm,
                &catch_apply,
                &[
                    method.clone(),
                    must_not_run.clone(),
                    Value::Null,
                    throwing_length_list.clone()
                ]
            )
            .as_f64(),
            123.0
        );
    }
    trace.lock().unwrap().clear();
    for method in [&apply, &reflect_apply] {
        assert_apply_error(
            &mut vm,
            &catch_apply,
            &get,
            method,
            &object_with(&[]),
            Value::Object(observed.clone()),
            "TypeError",
        );
    }
    assert!(trace.lock().unwrap().is_empty());
    vm.register_host_fn("benchmark", "returnNull", Box::new(|_, _| Value::Null));
    let returns_null = method_ref(&vm, "returnNull");
    let no_args = Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new())));
    assert!(matches!(
        invoke(
            &mut vm,
            &reflect_apply,
            &[returns_null, Value::Null, no_args]
        ),
        Value::Null
    ));
    assert_eq!(calls.load(Ordering::Relaxed), 0);
    let bind = host_ref(&vm, "ecma:function", "bind");
    let bound_sum = invoke(
        &mut vm,
        &bind,
        &[sum.clone(), Value::Null, Value::I32(1), Value::I32(2)],
    );
    assert_eq!(
        invoke(
            &mut vm,
            &call,
            &[bound_sum.clone(), Value::I32(99), Value::I32(3)]
        )
        .as_f64(),
        6.0
    );
    let final_arg = Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
        Value::I32(3),
    ])));
    for method in [&apply, &reflect_apply] {
        assert_eq!(
            invoke(
                &mut vm,
                method,
                &[bound_sum.clone(), Value::I32(99), final_arg.clone()]
            )
            .as_f64(),
            6.0
        );
    }
    vm.register_host_fn(
        "benchmark",
        "callableProxyApply",
        Box::new(|_, args| {
            assert_eq!(args.len(), 4);
            let Value::Object(arguments) = &args[3] else {
                panic!("proxy args must be an array")
            };
            let arguments = arguments.lock().unwrap();
            let ObjectKind::Array(values) = &arguments.kind else {
                panic!("proxy args must be array-kind")
            };
            Value::F64(args[2].as_f64() + values.iter().map(Value::as_f64).sum::<f64>())
        }),
    );
    let handler = object_with(&[("apply", method_ref(&vm, "callableProxyApply"))]);
    let callable_proxy = invoke(&mut vm, &proxy_new, &[sum, handler]);
    assert_eq!(
        invoke(
            &mut vm,
            &call,
            &[
                callable_proxy.clone(),
                Value::I32(99),
                Value::I32(5),
                Value::I32(7)
            ]
        )
        .as_f64(),
        111.0
    );
    let proxy_args = Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
        Value::I32(5),
        Value::I32(7),
    ])));
    for (label, method) in [
        ("Function.apply", &apply),
        ("Reflect.apply", &reflect_apply),
    ] {
        assert_eq!(
            invoke(
                &mut vm,
                method,
                &[callable_proxy.clone(), Value::I32(99), proxy_args.clone()]
            )
            .as_f64(),
            111.0,
            "{label} callable proxy"
        );
    }
    println!(
        "apply boundary, order, sparse, accessor, proxy, bound-function, coercion, invalid-list, abrupt-completion, and mutation checks passed"
    );
    if ecma::perf::enabled() {
        let samples = ecma::perf::snapshot();
        assert!(
            samples
                .iter()
                .any(|sample| sample.name.as_ref() == "ecma:object:get" && sample.calls > 0)
        );
        assert!(
            samples
                .iter()
                .any(|sample| sample.name.as_ref() == "ecma:function:apply" && sample.calls > 0)
        );
        assert!(
            samples
                .iter()
                .any(|sample| sample.name.as_ref() == "ecma:reflect:apply" && sample.calls > 0)
        );
        assert!(
            samples
                .iter()
                .all(|sample| sample.exclusive <= sample.inclusive)
        );
        println!("{}", ecma::perf::report());
        assert!(ecma::perf::reset());
        assert!(ecma::perf::snapshot().is_empty());
        let nested_get = get.clone();
        let nested_object = object_with(&[("__proto__", Value::Null), ("value", Value::I32(42))]);
        let nested_key = ecma::keys::string_value("value");
        vm.register_host_fn(
            "benchmark",
            "profileNested",
            Box::new(move |ctx, _| {
                // An active outer import must keep its accounting stack intact.
                assert!(!ecma::perf::reset());
                ctx.invoke(&nested_get, &[nested_object.clone(), nested_key.clone()])
            }),
        );
        let nested = method_ref(&vm, "profileNested");
        assert_eq!(
            invoke(&mut vm, &call, &[nested, Value::Null]).as_f64(),
            42.0
        );
        let nested_samples = ecma::perf::snapshot();
        let outer = nested_samples
            .iter()
            .find(|sample| sample.name.as_ref() == "ecma:function:call")
            .unwrap();
        let inner = nested_samples
            .iter()
            .find(|sample| sample.name.as_ref() == "ecma:object:get")
            .unwrap();
        assert_eq!(outer.calls, 1);
        assert_eq!(inner.calls, 1);
        assert!(outer.inclusive >= inner.inclusive);
        assert_eq!(
            outer.exclusive,
            outer.inclusive.saturating_sub(inner.inclusive)
        );
        assert!(ecma::perf::reset());
        assert!(ecma::perf::snapshot().is_empty());
        println!("profiling coverage, nested-time subtraction, and active-reset checks passed");
    }
}

fn assert_apply_error(
    vm: &mut VM,
    catcher: &Value,
    get: &Value,
    method: &Value,
    target: &Value,
    list: Value,
    expected_name: &str,
) {
    let error = invoke(
        vm,
        catcher,
        &[method.clone(), target.clone(), Value::Null, list],
    );
    let name = invoke(vm, get, &[error, ecma::keys::string_value("name")]);
    assert!(matches!(name, Value::String(text) if text.as_ref() == expected_name));
}
