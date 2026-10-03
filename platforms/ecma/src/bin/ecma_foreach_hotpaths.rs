//! Compiler-free entry-materialization and foreach-shaped import benchmark.
use indexmap::IndexMap;
use std::borrow::Cow;
use std::hint::black_box;
use std::time::Instant;
use vybe_platform_ecma as ecma;
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{TypeDef, VM, Value};

fn host_ref(vm: &VM, module: &str, name: &str) -> Value {
    let mut object = Object::new();
    object.kind = ObjectKind::HostFunction(vm.resolve_host_function_index(module, name).unwrap());
    Value::Object(vybe_runtime::heap::alloc(object))
}

fn invoke(vm: &mut VM, function: &Value, args: &[Value]) -> Value {
    vm.invoke(function, args)
        .expect("foreach benchmark invocation failed")
}

fn foreach_round(vm: &mut VM, entries: &Value, get: &Value, length: &Value, source: &Value) -> i32 {
    let pairs = invoke(vm, entries, &[source.clone()]);
    let count = invoke(vm, length, &[pairs.clone()]).as_i32();
    let mut sum = 0;
    for index in 0..count {
        let pair = invoke(vm, get, &[pairs.clone(), Value::I32(index)]);
        black_box(invoke(vm, get, &[pair.clone(), Value::I32(0)]));
        sum += invoke(vm, get, &[pair, Value::I32(1)]).as_i32();
    }
    sum
}

fn check_projection(
    vm: &mut VM,
    keys: &Value,
    values: &Value,
    get: &Value,
    length: &Value,
    source: &Value,
    expected_keys: &[&str],
    expected_values: &[i32],
) {
    let projected_keys = invoke(vm, keys, &[source.clone()]);
    let projected_values = invoke(vm, values, &[source.clone()]);
    assert_eq!(
        invoke(vm, length, &[projected_keys.clone()]).as_i32() as usize,
        expected_keys.len()
    );
    assert_eq!(
        invoke(vm, length, &[projected_values.clone()]).as_i32() as usize,
        expected_values.len()
    );
    for (index, (key, value)) in expected_keys.iter().zip(expected_values).enumerate() {
        let actual = invoke(vm, get, &[projected_keys.clone(), Value::I32(index as i32)]);
        assert!(matches!(actual, Value::String(text) if text.as_ref() == *key));
        assert_eq!(
            invoke(
                vm,
                get,
                &[projected_values.clone(), Value::I32(index as i32)]
            )
            .as_i32(),
            *value
        );
    }
}

fn main() {
    let iterations: usize = std::env::args()
        .nth(1)
        .map(|value| value.parse().unwrap())
        .unwrap_or(1000);
    assert!(iterations > 0);
    let mut vm = VM::new();
    ecma::register(&mut vm);
    let get = host_ref(&vm, "ecma:array", "get");
    let get_value = host_ref(&vm, "ecma:array", "getValue");
    let length = host_ref(&vm, "ecma:array", "length");
    let entries = host_ref(&vm, "ecma:object", "entries");
    let object_keys = host_ref(&vm, "ecma:object", "keys");
    let object_values = host_ref(&vm, "ecma:object", "values");
    let mut definition = TypeDef::new("ECMAEntriesNativeFixture");
    definition.add_field("nativeSlot");
    let type_id = vm.type_registry.register(definition);
    assert_ne!(type_id, 0);
    let mut typed = Object::new_typed_with_fields(type_id, 1);
    typed.fields[0] = Value::I32(123);
    typed
        .properties
        .insert("dynamicValue".into(), Value::I32(7));
    let mut typed_metadata = typed.clone();
    typed_metadata
        .properties
        .insert("__proto__".into(), Value::Null);
    let typed = Value::Object(vybe_runtime::heap::alloc(typed));
    let typed_metadata = Value::Object(vybe_runtime::heap::alloc(typed_metadata));
    // Preserve the established fallback result, without claiming the
    // property-bag adapter fully projects native class field descriptors.
    assert_eq!(
        foreach_round(&mut vm, &entries, &get, &length, &typed),
        foreach_round(&mut vm, &entries, &get, &length, &typed_metadata)
    );
    if let Value::Object(object) = &typed {
        let object = object.lock().unwrap();
        assert_eq!(object.type_id, type_id);
        assert_eq!(object.fields[0].as_i32(), 123);
    }
    let mut plain = Object::new();
    for index in 0..8 {
        plain
            .properties
            .insert(format!("field{index}"), Value::I32(index + 1));
    }
    let mut metadata = plain.clone();
    metadata.properties.insert("__proto__".into(), Value::Null);
    let plain = Value::Object(vybe_runtime::heap::alloc(plain));
    let metadata = Value::Object(vybe_runtime::heap::alloc(metadata));
    let expected_keys = [
        "field0", "field1", "field2", "field3", "field4", "field5", "field6", "field7",
    ];
    let expected_values = [1, 2, 3, 4, 5, 6, 7, 8];
    for (source_label, source) in [("plain", &plain), ("metadata fallback", &metadata)] {
        check_projection(
            &mut vm,
            &object_keys,
            &object_values,
            &get,
            &length,
            source,
            &expected_keys,
            &expected_values,
        );
        for (label, function) in [("keys", &object_keys), ("values", &object_values)] {
            let start = Instant::now();
            for _ in 0..iterations {
                black_box(invoke(&mut vm, function, &[source.clone()]));
            }
            println!(
                "ordinary {source_label} {label} (8 properties): {:.1} ns/call",
                start.elapsed().as_secs_f64() * 1e9 / iterations as f64
            );
        }
    }
    for (label, source) in [
        ("ordinary plain", &plain),
        ("ordinary metadata fallback", &metadata),
    ] {
        assert_eq!(foreach_round(&mut vm, &entries, &get, &length, source), 36);
        let start = Instant::now();
        for _ in 0..iterations {
            assert_eq!(
                black_box(foreach_round(&mut vm, &entries, &get, &length, source)),
                36
            );
        }
        println!(
            "{label} entries + foreach (8 pairs): {:.1} ns/round",
            start.elapsed().as_secs_f64() * 1e9 / iterations as f64
        );
    }
    let mut numeric = Object::new();
    numeric.properties.insert("10".into(), Value::I32(10));
    numeric.properties.insert("2".into(), Value::I32(2));
    numeric.properties.insert("tail".into(), Value::I32(7));
    let numeric = Value::Object(vybe_runtime::heap::alloc(numeric));
    check_projection(
        &mut vm,
        &object_keys,
        &object_values,
        &get,
        &length,
        &numeric,
        &["2", "10", "tail"],
        &[2, 10, 7],
    );
    let ordered = invoke(&mut vm, &entries, &[numeric]);
    for (index, expected) in ["2", "10", "tail"].iter().enumerate() {
        let pair = invoke(&mut vm, &get, &[ordered.clone(), Value::I32(index as i32)]);
        let key = invoke(&mut vm, &get, &[pair, Value::I32(0)]);
        assert!(matches!(key, Value::String(text) if text.as_ref() == *expected));
    }
    let mut descriptor = Object::new();
    descriptor
        .properties
        .insert("value".into(), Value::I32(999));
    descriptor
        .properties
        .insert("enumerable".into(), Value::Bool(false));
    descriptor
        .properties
        .insert("configurable".into(), Value::Bool(true));
    descriptor
        .properties
        .insert("writable".into(), Value::Bool(true));
    let define = host_ref(&vm, "ecma:object", "defineProperty");
    invoke(
        &mut vm,
        &define,
        &[
            plain.clone(),
            ecma::keys::string_value("hidden"),
            Value::Object(vybe_runtime::heap::alloc(descriptor)),
        ],
    );
    assert_eq!(foreach_round(&mut vm, &entries, &get, &length, &plain), 36);
    check_projection(
        &mut vm,
        &object_keys,
        &object_values,
        &get,
        &length,
        &plain,
        &expected_keys,
        &expected_values,
    );
    let mut map = IndexMap::new();
    for index in 0..8 {
        map.insert(
            ecma::keys::owned_string_value(format!("field{index}")),
            Value::I32(index + 1),
        );
    }
    map.insert(ecma::keys::string_value("0"), Value::I32(17));
    map.insert(Value::I32(7), Value::I32(29));
    map.insert(ecma::keys::string_value("true"), Value::I32(11));
    let mut object = Object::new();
    object.kind = ObjectKind::Map(map);
    let source = Value::Object(vybe_runtime::heap::alloc(object));
    let key = ecma::keys::string_value("field3");
    assert!(matches!(
        ecma::keys::map_lookup_key_cow(&key),
        Cow::Borrowed(_)
    ));
    assert!(matches!(
        ecma::keys::map_lookup_key_cow(&Value::Bool(true)),
        Cow::Owned(_)
    ));
    for function in [&get, &get_value] {
        assert_eq!(
            invoke(&mut vm, function, &[source.clone(), key.clone()]).as_i32(),
            4
        );
        assert_eq!(
            invoke(&mut vm, function, &[source.clone(), Value::I32(0)]).as_i32(),
            17
        );
        assert_eq!(
            invoke(
                &mut vm,
                function,
                &[source.clone(), ecma::keys::string_value("7")]
            )
            .as_i32(),
            29
        );
    }
    assert_eq!(
        invoke(&mut vm, &get_value, &[source.clone(), Value::Bool(true)]).as_i32(),
        11
    );
    let lookup_args = [source.clone(), key];
    println!(
        "foreach imports; setup excluded, VM dispatch included; profile={}",
        if cfg!(debug_assertions) {
            "debug"
        } else {
            "optimized"
        }
    );
    for _ in 0..iterations.min(100) {
        black_box(invoke(&mut vm, &get, &lookup_args));
    }
    let start = Instant::now();
    for _ in 0..iterations {
        black_box(invoke(&mut vm, &get, black_box(&lookup_args)));
    }
    println!(
        "array.get.map-string: {:.1} ns/call",
        start.elapsed().as_secs_f64() * 1e9 / iterations as f64
    );
    assert_eq!(foreach_round(&mut vm, &entries, &get, &length, &source), 93);
    let start = Instant::now();
    for _ in 0..iterations {
        assert_eq!(
            black_box(foreach_round(&mut vm, &entries, &get, &length, &source)),
            93
        );
    }
    println!(
        "entries + foreach (11 pairs): {:.1} ns/round",
        start.elapsed().as_secs_f64() * 1e9 / iterations as f64
    );
    if ecma::perf::enabled() {
        let samples = ecma::perf::snapshot();
        for name in [
            "ecma:object:keys",
            "ecma:object:values",
            "ecma:object:entries",
            "ecma:array:length",
            "ecma:array:get",
        ] {
            assert!(
                samples
                    .iter()
                    .any(|sample| sample.name.as_ref() == name && sample.calls > 0),
                "missing {name}"
            );
        }
        println!("{}", ecma::perf::report());
    }
    println!(
        "native type/slot fallback, borrowed-key, numeric ordering/fallback, non-enumerable, coercion, foreach checksum and profiling checks passed"
    );
}
