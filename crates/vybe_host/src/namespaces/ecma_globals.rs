//! New ECMA-262 global constructors and namespaces.
//!
//! Wires up `Symbol`, `Reflect`, `Atomics`, `BigInt`, `Iterator`, plus
//! `Math` Stage-3 accumulator extensions (`minOf`, `maxOf`,
//! `sumPrecise`) and the proper `globalThis` singleton — all so JS code
//! can do `Symbol.iterator`, `Reflect.apply(...)`, `Atomics.add(...)`,
//! `BigInt(123n)`, etc., the same way real JS runtimes expose them.
//!
//! Each `Foo` global is an Object whose properties are HostFunction refs
//! pointing at the registered `ecma:foo:*` host fns. Static methods
//! (`Symbol.for`, `Reflect.apply`) become properties; constructors
//! (`Symbol`, `BigInt`) double as the namespace object's own
//! invocation target via the `__call` convention.

use super::*;

pub fn register(vm: &mut VM) {
    // ── Object / boxed primitive constructors ─────────────────────
    let object = host_fn_ref(vm, "ecma:object", "Object");
    set_prop(&object, "name", Value::String(Arc::from("Object")));
    let object_proto = crate::ecma::object::shared_object_prototype();
    set_prop(&object_proto, "constructor", object.clone());
    for name in &["toString", "toLocaleString", "valueOf", "hasOwnProperty", "propertyIsEnumerable", "isPrototypeOf"] {
        let idx = *vm.host_registry.get(&("ecma:object".to_string(), (*name).to_string())).expect("ecma:object prototype method must be registered");
        set_prop(&object_proto, name, receiver_host_fn_ref("ecma:object", name, idx));
    }
    set_prop(&object, "prototype", object_proto.clone());
    for name in &[
        "keys", "values", "entries", "assign", "freeze", "fromEntries", "hasOwn",
        "create", "seal", "isFrozen", "isSealed", "is", "getPrototypeOf",
        "getOwnPropertyNames", "defineProperty",
    ] {
        set_prop(&object, name, host_fn_ref(vm, "ecma:object", name));
    }
    set_prop(&object, "groupBy", Value::Bool(true));
    vm.globals.insert("Object".to_string(), object.clone());
    vm.globals.insert("object".to_string(), object.clone());

    let number = host_fn_ref(vm, "ecma:number", "Number");
    set_prop(&number, "name", Value::String(Arc::from("Number")));
    let number_proto = crate::ecma::number::shared_number_prototype();
    set_prop(&number_proto, "constructor", number.clone());
    set_prop(&number_proto, "__proto__", object_proto.clone());
    for name in &["toString", "valueOf", "toFixed", "toExponential", "toPrecision", "toLocaleString"] {
        let idx = *vm.host_registry.get(&("ecma:number".to_string(), (*name).to_string())).expect("ecma:number prototype method must be registered");
        set_prop(&number_proto, name, receiver_host_fn_ref("ecma:number", name, idx));
    }
    set_prop(&number, "prototype", number_proto);
    for name in &["parseInt", "parseFloat", "isNaN", "isFinite", "isInteger", "isSafeInteger"] {
        set_prop(&number, name, host_fn_ref(vm, "ecma:number", name));
    }
    for (name, export) in &[
        ("MAX_SAFE_INTEGER", "MAX_SAFE_INTEGER"),
        ("MIN_SAFE_INTEGER", "MIN_SAFE_INTEGER"),
        ("MAX_VALUE", "MAX_VALUE"),
        ("MIN_VALUE", "MIN_VALUE"),
        ("EPSILON", "EPSILON"),
        ("POSITIVE_INFINITY", "POSITIVE_INFINITY"),
        ("NEGATIVE_INFINITY", "NEGATIVE_INFINITY"),
        ("NaN", "NaN"),
    ] {
        set_prop(&number, name, host_fn_ref(vm, "ecma:number", export));
    }
    vm.globals.insert("Number".to_string(), number.clone());
    vm.globals.insert("number".to_string(), number.clone());

    let string = host_fn_ref(vm, "ecma:string", "String");
    set_prop(&string, "name", Value::String(Arc::from("String")));
    let string_proto = crate::ecma::string::shared_string_prototype();
    set_prop(&string_proto, "constructor", string.clone());
    set_prop(&string_proto, "__proto__", object_proto.clone());
    for name in &[
        "toString", "valueOf", "length", "charAt", "charCodeAt", "codePointAt", "at",
        "concat", "substring", "slice", "toUpperCase", "toLowerCase", "toLocaleUpperCase",
        "toLocaleLowerCase", "trim", "trimStart", "trimEnd", "padStart", "padEnd",
        "includes", "indexOf", "lastIndexOf", "startsWith", "endsWith", "repeat",
        "replace", "replaceAll", "split", "localeCompare", "normalize",
    ] {
        let idx = *vm.host_registry.get(&("ecma:string".to_string(), (*name).to_string())).expect("ecma:string prototype method must be registered");
        set_prop(&string_proto, name, receiver_host_fn_ref("ecma:string", name, idx));
    }
    set_prop(&string, "prototype", string_proto);
    for name in &["fromCharCode", "fromCodePoint", "raw"] {
        set_prop(&string, name, host_fn_ref(vm, "ecma:string", name));
    }
    vm.globals.insert("String".to_string(), string.clone());
    vm.globals.insert("string".to_string(), string.clone());

    let boolean = host_fn_ref(vm, "ecma:boolean", "Boolean");
    set_prop(&boolean, "name", Value::String(Arc::from("Boolean")));
    let boolean_proto = crate::ecma::boolean::shared_boolean_prototype();
    set_prop(&boolean_proto, "constructor", boolean.clone());
    set_prop(&boolean_proto, "__proto__", object_proto.clone());
    let boolean_to_string = *vm.host_registry.get(&("ecma:boolean".to_string(), "toString".to_string())).expect("ecma:boolean.toString must be registered");
    let boolean_value_of = *vm.host_registry.get(&("ecma:boolean".to_string(), "valueOf".to_string())).expect("ecma:boolean.valueOf must be registered");
    set_prop(&boolean_proto, "toString", receiver_host_fn_ref("ecma:boolean", "toString", boolean_to_string));
    set_prop(&boolean_proto, "valueOf", receiver_host_fn_ref("ecma:boolean", "valueOf", boolean_value_of));
    set_prop(&boolean, "prototype", boolean_proto);
    vm.globals.insert("Boolean".to_string(), boolean.clone());
    vm.globals.insert("boolean".to_string(), boolean.clone());

    let function = Value::Object(Arc::new(Mutex::new(Object::new())));
    set_prop(&function, "name", Value::String(Arc::from("Function")));
    let function_proto = crate::ecma::function::shared_function_prototype();
    set_prop(&function_proto, "constructor", function.clone());
    set_prop(&function_proto, "__proto__", object_proto.clone());
    for name in &["bind", "call", "apply"] {
        let idx = *vm.host_registry.get(&("ecma:function".to_string(), (*name).to_string())).expect("ecma:function prototype method must be registered");
        set_prop(&function_proto, name, receiver_host_fn_ref("ecma:function", name, idx));
    }
    set_prop(&function, "prototype", function_proto.clone());
    set_prop(&function, "__proto__", function_proto);
    vm.globals.insert("Function".to_string(), function.clone());
    vm.globals.insert("function".to_string(), function.clone());

    let array = host_fn_ref(vm, "ecma:array", "new");
    set_prop(&array, "name", Value::String(Arc::from("Array")));
    set_prop(&array, "__proto__", crate::ecma::function::shared_function_prototype());
    let array_proto = crate::ecma::array::shared_array_prototype();
    set_prop(&array_proto, "constructor", array.clone());
    set_prop(&array_proto, "__proto__", object_proto.clone());
    for name in &[
        "at", "concat", "copyWithin", "entries", "every", "fill", "filter", "find",
        "findIndex", "findLast", "findLastIndex", "flat", "flatMap", "forEach", "group",
        "groupToMap", "includes", "indexOf", "join", "lastIndexOf", "map", "pop", "push",
        "reduce", "reduceRight", "reverse", "shift", "slice", "some", "sort", "splice",
        "toReversed", "toSorted", "toSpliced", "unshift", "values",
    ] {
        let idx = *vm.host_registry.get(&("ecma:array".to_string(), (*name).to_string())).expect("ecma:array prototype method must be registered");
        set_prop(&array_proto, name, receiver_host_fn_ref("ecma:array", name, idx));
    }
    set_prop(&array, "prototype", array_proto);
    for name in &["from", "fromAsync", "isArray", "of"] {
        set_prop(&array, name, host_fn_ref(vm, "ecma:array", name));
    }
    vm.globals.insert("Array".to_string(), array.clone());
    vm.globals.insert("array".to_string(), array.clone());

    // ── Symbol ─────────────────────────────────────────────────────
    // `Symbol` is callable as `Symbol(desc)` AND has static properties
    // (`Symbol.for`, `Symbol.iterator`, etc.). Build a Value that is
    // both a HostFunction (kind=HostFunction(idx)) and carries the
    // static properties as Object properties.
    let sym = host_fn_with_props(vm, "ecma:symbol", "Symbol");
    set_prop(&sym, "for",            host_fn_ref(vm, "ecma:symbol", "for"));
    set_prop(&sym, "keyFor",         host_fn_ref(vm, "ecma:symbol", "keyFor"));
    // Well-knowns — materialized as Symbol primitives directly. The
    // sentinel strings are the same that `ecma:symbol` exposes via
    // its 0-arg getters (kept in sync; both paths share the convention
    // that well-known symbols stringify to "@@<name>").
    for name in &[
        "iterator", "asyncIterator", "toPrimitive", "hasInstance",
        "toStringTag", "isConcatSpreadable", "unscopables", "match",
        "matchAll", "replace", "search", "split", "species", "dispose",
        "asyncDispose",
    ] {
        let sentinel = format!("@@{}", name);
        set_prop(&sym, name, Value::Symbol(Arc::from(sentinel.as_str())));
    }
    vm.globals.insert("Symbol".to_string(), sym);
    vm.globals.insert("symbol".to_string(), vm.globals.get("Symbol").cloned().unwrap_or(Value::Null));

    // ── Reflect ────────────────────────────────────────────────────
    let reflect = ensure_namespace(vm, &["Reflect"]);
    for name in &[
        "apply", "construct", "get", "set", "has", "deleteProperty",
        "ownKeys", "getOwnPropertyDescriptor", "defineProperty",
        "getPrototypeOf", "setPrototypeOf", "isExtensible",
        "preventExtensions",
    ] {
        set_prop(&reflect, name, host_fn_ref(vm, "ecma:reflect", name));
    }

    // ── Atomics ────────────────────────────────────────────────────
    let atomics = ensure_namespace(vm, &["Atomics"]);
    for name in &[
        "add", "sub", "and", "or", "xor", "exchange", "compareExchange",
        "load", "store", "isLockFree", "wait", "notify",
    ] {
        set_prop(&atomics, name, host_fn_ref(vm, "ecma:atomics", name));
    }

    // ── BigInt ─────────────────────────────────────────────────────
    // Same callable-with-static-properties shape as Symbol.
    let bigint = host_fn_with_props(vm, "ecma:bigint", "BigInt");
    set_prop(&bigint, "asIntN",  host_fn_ref(vm, "ecma:bigint", "asIntN"));
    set_prop(&bigint, "asUintN", host_fn_ref(vm, "ecma:bigint", "asUintN"));
    vm.globals.insert("BigInt".to_string(), bigint);
    vm.globals.insert("bigint".to_string(), vm.globals.get("BigInt").cloned().unwrap_or(Value::Null));

    // ── Iterator (Stage-3 helpers) ─────────────────────────────────
    let iter = ensure_namespace(vm, &["Iterator"]);
    set_prop(&iter, "from",  host_fn_ref(vm, "ecma:iterator", "from"));
    set_prop(&iter, "range", host_fn_ref(vm, "ecma:iterator", "range"));

    // ── Math Stage-3 accumulators ──────────────────────────────────
    // Math global already created by namespaces/math.rs; just add the new
    // properties so they show up alongside the existing ones.
    let math = ensure_namespace(vm, &["Math"]);
    set_prop(&math, "minOf",       host_fn_ref(vm, "ecma:math", "minOf"));
    set_prop(&math, "maxOf",       host_fn_ref(vm, "ecma:math", "maxOf"));
    set_prop(&math, "sumPrecise",  host_fn_ref(vm, "ecma:math", "sumPrecise"));

    // ── TypedArrays — callable constructors + static properties ──────────
    // Each variant is a callable host function (for `new Int8Array(...)`)
    // AND carries static properties (`from`, `of`, `BYTES_PER_ELEMENT`).
    const TYPED_ARRAY_GLOBALS: &[(&str, &str, i32)] = &[
        ("Int8Array",         "ecma:int8array",     1),
        ("Uint8Array",        "ecma:uint8array",    1),
        ("Uint8ClampedArray", "ecma:uint8clamped",  1),
        ("Int16Array",        "ecma:int16array",    2),
        ("Uint16Array",       "ecma:uint16array",   2),
        ("Int32Array",        "ecma:int32array",    4),
        ("Uint32Array",       "ecma:uint32array",   4),
        ("Float32Array",      "ecma:float32array",  4),
        ("Float64Array",      "ecma:float64array",  8),
        ("BigInt64Array",     "ecma:bigint64array", 8),
        ("BigUint64Array",    "ecma:biguint64array",8),
    ];
    for (global_name, module, bpe) in TYPED_ARRAY_GLOBALS {
        let ctor = host_fn_ref(vm, module, "new");
        if !matches!(ctor, Value::Null) {
            set_prop(&ctor, "from",              host_fn_ref(vm, module, "from"));
            set_prop(&ctor, "of",                host_fn_ref(vm, module, "of"));
            set_prop(&ctor, "BYTES_PER_ELEMENT", Value::I32(*bpe));
            vm.globals.insert(global_name.to_string(), ctor);
        }
    }

    // ── globalThis — proper §19.3.1 singleton ──────────────────────
    // Pulls the shared process-global Object that `ecma:globalThis.get`
    // also returns, so identity holds across both access patterns.
    vm.globals.insert("globalThis".to_string(), crate::ecma::global_this::shared_singleton());
}

/// Build a Value that is callable as a host function AND can carry
/// static properties — like real JS where `Symbol` and `BigInt` are
/// both callable AND have static methods (`Symbol.for`, `BigInt.asIntN`).
fn host_fn_with_props(vm: &VM, module: &str, name: &str) -> Value {
    host_fn_ref(vm, module, name)
}
