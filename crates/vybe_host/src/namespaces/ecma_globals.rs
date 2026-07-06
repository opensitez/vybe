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
    set_constructor_once(&object_proto, object.clone());
    if let Value::Object(proto) = &object_proto {
        crate::ecma::object::track_nonenum(proto, "constructor");
        crate::ecma::object::track_nonenum(proto, "constructor");
    }
    // §20.1.3: the values stored ON %Object.prototype% are the RAW
    // intrinsics. A borrowed `Object.prototype.hasOwnProperty.call(o, k)`
    // must NOT consult `o`'s own override — override dispatch belongs to
    // ordinary property lookup on the receiver, which finds the own
    // method before ever reaching the prototype (the override-checking
    // `ecma:object.hasOwnProperty` exists for the compiler's direct
    // value_methods dispatch, where no runtime lookup happens).
    for (name, registry_name) in &[
        ("toString", "toString"),
        ("toLocaleString", "toLocaleString"),
        ("valueOf", "valueOf"),
        ("hasOwnProperty", "hasOwnPropertyIntrinsic"),
        ("propertyIsEnumerable", "propertyIsEnumerable"),
        ("isPrototypeOf", "isPrototypeOf"),
    ] {
        let idx = *vm
            .host_registry
            .get(&("ecma:object".to_string(), (*registry_name).to_string()))
            .expect("ecma:object prototype method must be registered");
        set_prop(
            &object_proto,
            name,
            receiver_host_fn_ref("ecma:object", name, idx),
        );
        if let Value::Object(proto) = &object_proto {
            crate::ecma::object::track_nonenum(proto, name);
            let lower = name.to_lowercase();
            if lower != *name {
                crate::ecma::object::track_nonenum(proto, &lower);
            }
        }
    }
    set_prop(&object, "prototype", object_proto.clone());
    for name in &[
        "keys",
        "values",
        "entries",
        "assign",
        "freeze",
        "fromEntries",
        "hasOwn",
        "create",
        "seal",
        "isFrozen",
        "isSealed",
        "is",
        "getPrototypeOf",
        "getOwnPropertyNames",
        "defineProperty",
    ] {
        set_prop(&object, name, host_fn_ref(vm, "ecma:object", name));
    }
    set_prop(&object, "groupBy", Value::Bool(true));
    vm.globals.insert("Object".to_string(), object.clone());
    vm.globals.insert("object".to_string(), object.clone());

    let number = host_fn_ref(vm, "ecma:number", "Number");
    set_prop(&number, "name", Value::String(Arc::from("Number")));
    let number_proto = crate::ecma::number::shared_number_prototype();
    set_constructor_once(&number_proto, number.clone());
    set_prop(&number_proto, "__proto__", object_proto.clone());
    for name in &[
        "toString",
        "valueOf",
        "toFixed",
        "toExponential",
        "toPrecision",
        "toLocaleString",
    ] {
        let idx = *vm
            .host_registry
            .get(&("ecma:number".to_string(), (*name).to_string()))
            .expect("ecma:number prototype method must be registered");
        set_prop(
            &number_proto,
            name,
            receiver_host_fn_ref("ecma:number", name, idx),
        );
    }
    set_prop(&number, "prototype", number_proto);
    for name in &[
        "parseInt",
        "parseFloat",
        "isNaN",
        "isFinite",
        "isInteger",
        "isSafeInteger",
    ] {
        set_prop(&number, name, host_fn_ref(vm, "ecma:number", name));
    }
    set_prop(&number, "MAX_SAFE_INTEGER", Value::F64(9007199254740991.0));
    set_prop(&number, "MIN_SAFE_INTEGER", Value::F64(-9007199254740991.0));
    set_prop(&number, "MAX_VALUE", Value::F64(f64::MAX));
    set_prop(&number, "MIN_VALUE", Value::F64(f64::from_bits(1)));
    set_prop(&number, "EPSILON", Value::F64(f64::EPSILON));
    set_prop(&number, "POSITIVE_INFINITY", Value::F64(f64::INFINITY));
    set_prop(&number, "NEGATIVE_INFINITY", Value::F64(f64::NEG_INFINITY));
    set_prop(&number, "NaN", Value::F64(f64::NAN));
    vm.globals.insert("Number".to_string(), number.clone());
    vm.globals.insert("number".to_string(), number.clone());

    let string = host_fn_ref(vm, "ecma:string", "String");
    set_prop(&string, "name", Value::String(Arc::from("String")));
    let string_proto = crate::ecma::string::shared_string_prototype();
    set_constructor_once(&string_proto, string.clone());
    set_prop(&string_proto, "__proto__", object_proto.clone());
    for name in &[
        "toString",
        "valueOf",
        "length",
        "charAt",
        "charCodeAt",
        "codePointAt",
        "at",
        "concat",
        "substring",
        "slice",
        "toUpperCase",
        "toLowerCase",
        "toLocaleUpperCase",
        "toLocaleLowerCase",
        "trim",
        "trimStart",
        "trimEnd",
        "padStart",
        "padEnd",
        "includes",
        "indexOf",
        "lastIndexOf",
        "startsWith",
        "endsWith",
        "repeat",
        "replace",
        "replaceAll",
        "split",
        "localeCompare",
        "normalize",
    ] {
        let idx = *vm
            .host_registry
            .get(&("ecma:string".to_string(), (*name).to_string()))
            .expect("ecma:string prototype method must be registered");
        set_prop(
            &string_proto,
            name,
            receiver_host_fn_ref("ecma:string", name, idx),
        );
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
    set_constructor_once(&boolean_proto, boolean.clone());
    set_prop(&boolean_proto, "__proto__", object_proto.clone());
    let boolean_to_string = *vm
        .host_registry
        .get(&("ecma:boolean".to_string(), "toString".to_string()))
        .expect("ecma:boolean.toString must be registered");
    let boolean_value_of = *vm
        .host_registry
        .get(&("ecma:boolean".to_string(), "valueOf".to_string()))
        .expect("ecma:boolean.valueOf must be registered");
    set_prop(
        &boolean_proto,
        "toString",
        receiver_host_fn_ref("ecma:boolean", "toString", boolean_to_string),
    );
    set_prop(
        &boolean_proto,
        "valueOf",
        receiver_host_fn_ref("ecma:boolean", "valueOf", boolean_value_of),
    );
    set_prop(&boolean, "prototype", boolean_proto);
    vm.globals.insert("Boolean".to_string(), boolean.clone());
    vm.globals.insert("boolean".to_string(), boolean.clone());

    let function = Value::Object(Arc::new(Mutex::new(Object::new())));
    set_prop(&function, "name", Value::String(Arc::from("Function")));
    let function_proto = crate::ecma::function::shared_function_prototype();
    set_constructor_once(&function_proto, function.clone());
    set_prop(&function_proto, "__proto__", object_proto.clone());
    for name in &["bind", "call", "apply"] {
        let idx = *vm
            .host_registry
            .get(&("ecma:function".to_string(), (*name).to_string()))
            .expect("ecma:function prototype method must be registered");
        set_prop(
            &function_proto,
            name,
            receiver_host_fn_ref("ecma:function", name, idx),
        );
    }
    set_prop(&function, "prototype", function_proto.clone());
    set_prop(&function, "__proto__", function_proto);
    vm.globals.insert("Function".to_string(), function.clone());
    vm.globals.insert("function".to_string(), function.clone());

    let array = host_fn_ref(vm, "ecma:array", "new");
    set_prop(&array, "name", Value::String(Arc::from("Array")));
    set_prop(
        &array,
        "__proto__",
        crate::ecma::function::shared_function_prototype(),
    );
    let array_proto = crate::ecma::array::shared_array_prototype();
    set_constructor_once(&array_proto, array.clone());
    if let Value::Object(proto) = &array_proto {
        crate::ecma::object::track_nonenum(proto, "constructor");
    }
    set_prop(&array_proto, "__proto__", object_proto.clone());
    for name in &[
        "at",
        "concat",
        "copyWithin",
        "entries",
        "every",
        "fill",
        "filter",
        "find",
        "findIndex",
        "findLast",
        "findLastIndex",
        "flat",
        "flatMap",
        "forEach",
        "group",
        "groupToMap",
        "includes",
        "indexOf",
        "join",
        "lastIndexOf",
        "map",
        "pop",
        "push",
        "reduce",
        "reduceRight",
        "reverse",
        "shift",
        "slice",
        "some",
        "sort",
        "splice",
        "toReversed",
        "toSorted",
        "toSpliced",
        "unshift",
        "values",
    ] {
        let idx = *vm
            .host_registry
            .get(&("ecma:array".to_string(), (*name).to_string()))
            .expect("ecma:array prototype method must be registered");
        set_prop(
            &array_proto,
            name,
            receiver_host_fn_ref("ecma:array", name, idx),
        );
        if let Value::Object(proto) = &array_proto {
            crate::ecma::object::track_nonenum(proto, name);
            let lower = name.to_lowercase();
            if lower != *name {
                crate::ecma::object::track_nonenum(proto, &lower);
            }
        }
    }
    if let Some(idx) = vm
        .host_registry
        .get(&("ecma:array".to_string(), "values".to_string()))
        .copied()
    {
        set_prop(
            &array_proto,
            "iterator",
            receiver_host_fn_ref("ecma:array", "values", idx),
        );
        if let Value::Object(proto) = &array_proto {
            crate::ecma::object::track_nonenum(proto, "iterator");
            crate::ecma::object::track_nonenum(proto, "iterator");
        }
    }
    set_prop(&array, "prototype", array_proto);
    for name in &["from", "fromAsync", "isArray", "of"] {
        set_prop(&array, name, host_fn_ref(vm, "ecma:array", name));
    }
    vm.globals.insert("Array".to_string(), array.clone());
    vm.globals.insert("array".to_string(), array.clone());

    let date = host_fn_ref(vm, "ecma:date", "new");
    if !matches!(date, Value::Null) {
        set_prop(&date, "name", Value::String(Arc::from("Date")));
        set_prop(
            &date,
            "__proto__",
            crate::ecma::function::shared_function_prototype(),
        );
        let date_proto = crate::ecma::date::shared_date_prototype();
        set_prop(&date_proto, "constructor", date.clone());
        set_prop(&date_proto, "__proto__", object_proto.clone());
        for name in &[
            "getFullYear",
            "getYear",
            "getMonth",
            "getDate",
            "getDay",
            "getHours",
            "getMinutes",
            "getSeconds",
            "getMilliseconds",
            "getUTCFullYear",
            "getUTCMonth",
            "getUTCDate",
            "getUTCDay",
            "getUTCHours",
            "getUTCMinutes",
            "getUTCSeconds",
            "getUTCMilliseconds",
            "getTime",
            "getTimezoneOffset",
            "valueOf",
            "setTime",
            "setFullYear",
            "setMonth",
            "setDate",
            "setHours",
            "setMinutes",
            "setSeconds",
            "setMilliseconds",
            "setUTCFullYear",
            "setUTCMonth",
            "setUTCDate",
            "setUTCHours",
            "setUTCMinutes",
            "setUTCSeconds",
            "setUTCMilliseconds",
            "toISOString",
            "toString",
            "toUTCString",
            "toDateString",
            "toTimeString",
            "toJSON",
        ] {
            if let Some(idx) = vm
                .host_registry
                .get(&("ecma:date".to_string(), (*name).to_string()))
                .copied()
            {
                set_prop(
                    &date_proto,
                    name,
                    receiver_host_fn_ref("ecma:date", name, idx),
                );
            }
        }
        set_prop(&date, "prototype", date_proto);
        for name in &["now", "parse", "UTC"] {
            set_prop(&date, name, host_fn_ref(vm, "ecma:date", name));
        }
        vm.globals.insert("Date".to_string(), date.clone());
        vm.globals.insert("date".to_string(), date.clone());
    }

    // ── Symbol ─────────────────────────────────────────────────────
    // `Symbol` is callable as `Symbol(desc)` AND has static properties
    // (`Symbol.for`, `Symbol.iterator`, etc.). Build a Value that is
    // both a HostFunction (kind=HostFunction(idx)) and carries the
    // static properties as Object properties.
    let sym = host_fn_with_props(vm, "ecma:symbol", "Symbol");
    set_prop(&sym, "for", host_fn_ref(vm, "ecma:symbol", "for"));
    set_prop(&sym, "keyFor", host_fn_ref(vm, "ecma:symbol", "keyFor"));
    // Well-knowns — materialized as Symbol primitives directly. The
    // sentinel strings are the same that `ecma:symbol` exposes via
    // its 0-arg getters (kept in sync; both paths share the convention
    // that well-known symbols stringify to "@@<name>").
    for name in &[
        "iterator",
        "asyncIterator",
        "toPrimitive",
        "hasInstance",
        "toStringTag",
        "isConcatSpreadable",
        "unscopables",
        "match",
        "matchAll",
        "replace",
        "search",
        "split",
        "species",
        "dispose",
        "asyncDispose",
    ] {
        let sentinel = format!("@@{}", name);
        set_prop(&sym, name, Value::Symbol(Arc::from(sentinel.as_str())));
    }
    vm.globals.insert("Symbol".to_string(), sym);
    vm.globals.insert(
        "symbol".to_string(),
        vm.globals.get("Symbol").cloned().unwrap_or(Value::Null),
    );

    // ── Reflect ────────────────────────────────────────────────────
    let reflect = ensure_namespace(vm, &["Reflect"]);
    for name in &[
        "apply",
        "construct",
        "get",
        "set",
        "has",
        "deleteProperty",
        "ownKeys",
        "getOwnPropertyDescriptor",
        "defineProperty",
        "getPrototypeOf",
        "setPrototypeOf",
        "isExtensible",
        "preventExtensions",
    ] {
        set_prop(&reflect, name, host_fn_ref(vm, "ecma:reflect", name));
    }

    // ── Proxy ──────────────────────────────────────────────────────
    let proxy = ensure_namespace(vm, &["Proxy"]);
    set_prop(
        &proxy,
        "revocable",
        host_fn_ref(vm, "ecma:reflect", "proxyRevocable"),
    );

    // ── Atomics ────────────────────────────────────────────────────
    let atomics = ensure_namespace(vm, &["Atomics"]);
    for name in &[
        "add",
        "sub",
        "and",
        "or",
        "xor",
        "exchange",
        "compareExchange",
        "load",
        "store",
        "isLockFree",
        "wait",
        "notify",
    ] {
        set_prop(&atomics, name, host_fn_ref(vm, "ecma:atomics", name));
    }

    // ── BigInt ─────────────────────────────────────────────────────
    // Same callable-with-static-properties shape as Symbol.
    let bigint = host_fn_with_props(vm, "ecma:bigint", "BigInt");
    set_prop(&bigint, "asIntN", host_fn_ref(vm, "ecma:bigint", "asIntN"));
    set_prop(
        &bigint,
        "asUintN",
        host_fn_ref(vm, "ecma:bigint", "asUintN"),
    );
    vm.globals.insert("BigInt".to_string(), bigint);
    vm.globals.insert(
        "bigint".to_string(),
        vm.globals.get("BigInt").cloned().unwrap_or(Value::Null),
    );

    for name in &[
        "encodeURI",
        "decodeURI",
        "encodeURIComponent",
        "decodeURIComponent",
        "escape",
        "unescape",
        "btoa",
        "atob",
    ] {
        vm.globals
            .insert((*name).to_string(), host_fn_ref(vm, "ecma:string", name));
    }

    // ── Iterator (Stage-3 helpers) ─────────────────────────────────
    let iter = ensure_namespace(vm, &["Iterator"]);
    set_prop(&iter, "from", host_fn_ref(vm, "ecma:iterator", "from"));
    set_prop(&iter, "range", host_fn_ref(vm, "ecma:iterator", "range"));
    let async_iter = ensure_namespace(vm, &["AsyncIterator"]);
    set_prop(
        &async_iter,
        "from",
        host_fn_ref(vm, "ecma:iterator", "asyncFrom"),
    );

    // ── Math Stage-3 accumulators ──────────────────────────────────
    // Math global already created by namespaces/math.rs; just add the new
    // properties so they show up alongside the existing ones.
    let math = ensure_namespace(vm, &["Math"]);
    set_prop(&math, "minOf", host_fn_ref(vm, "ecma:math", "minOf"));
    set_prop(&math, "maxOf", host_fn_ref(vm, "ecma:math", "maxOf"));
    set_prop(
        &math,
        "sumPrecise",
        host_fn_ref(vm, "ecma:math", "sumPrecise"),
    );

    // ── TypedArrays — callable constructors + static properties ──────────
    // Each variant is a callable host function (for `new Int8Array(...)`)
    // AND carries static properties (`from`, `of`, `BYTES_PER_ELEMENT`).
    const TYPED_ARRAY_GLOBALS: &[(&str, &str, i32)] = &[
        ("Int8Array", "ecma:int8array", 1),
        ("Uint8Array", "ecma:uint8array", 1),
        ("Uint8ClampedArray", "ecma:uint8clamped", 1),
        ("Int16Array", "ecma:int16array", 2),
        ("Uint16Array", "ecma:uint16array", 2),
        ("Int32Array", "ecma:int32array", 4),
        ("Uint32Array", "ecma:uint32array", 4),
        ("Float32Array", "ecma:float32array", 4),
        ("Float64Array", "ecma:float64array", 8),
        ("BigInt64Array", "ecma:bigint64array", 8),
        ("BigUint64Array", "ecma:biguint64array", 8),
    ];
    for (global_name, module, bpe) in TYPED_ARRAY_GLOBALS {
        let ctor = host_fn_ref(vm, module, "new");
        if !matches!(ctor, Value::Null) {
            set_prop(&ctor, "name", Value::String(Arc::from(*global_name)));
            set_prop(&ctor, "from", host_fn_ref(vm, module, "from"));
            set_prop(&ctor, "of", host_fn_ref(vm, module, "of"));
            set_prop(&ctor, "BYTES_PER_ELEMENT", Value::I32(*bpe));
            vm.globals.insert(global_name.to_string(), ctor);
        }
    }

    // ── RegExp — constructor + static helpers ──────────────────────────
    // `new RegExp(...)` routes to ecma:regexp.new (see js/profile).
    // The constructor object also carries a static `test(pattern, str)`
    // convenience method for functional use: `RegExp.test(/\d+/, str)`.
    let regexp = host_fn_ref(vm, "ecma:regexp", "new");
    if !matches!(regexp, Value::Null) {
        set_prop(&regexp, "name", Value::String(Arc::from("RegExp")));
        set_prop(&regexp, "test", host_fn_ref(vm, "ecma:regexp", "test"));
        // §22.2.6: %RegExp.prototype% — instances link to it via
        // `__proto__` (stamped in ecma:regexp.new), so getPrototypeOf /
        // isPrototypeOf identity holds. The prototype carries the
        // receiver-shaped instance methods (the registered host fns take
        // the regex as arg 0, exactly what receiver_host_fn_ref prepends).
        let regexp_proto = crate::ecma::regexp::shared_regexp_prototype();
        set_constructor_once(&regexp_proto, regexp.clone());
        if let Value::Object(p) = &regexp_proto {
            crate::ecma::object::track_nonenum(p, "constructor");
        }
        for name in &["exec", "test", "toString"] {
            let idx = *vm
                .host_registry
                .get(&("ecma:regexp".to_string(), (*name).to_string()))
                .expect("ecma:regexp prototype method must be registered");
            set_prop(
                &regexp_proto,
                name,
                receiver_host_fn_ref("ecma:regexp", name, idx),
            );
            if let Value::Object(p) = &regexp_proto {
                crate::ecma::object::track_nonenum(p, name);
            }
        }
        set_prop(&regexp, "prototype", regexp_proto);
        vm.globals.insert("RegExp".to_string(), regexp.clone());
        vm.globals.insert("regexp".to_string(), regexp);
    }

    // ── globalThis — proper §19.3.1 singleton ──────────────────────
    // Pulls the shared process-global Object that `ecma:globalThis.get`
    // also returns, so identity holds across both access patterns.
    vm.globals.insert(
        "globalThis".to_string(),
        crate::ecma::global_this::shared_singleton(),
    );

    // ── Canonical constructor anchors (`__ctor_<Name>`) ────────────────
    // The user-facing constructor globals (`Array`, `Object`, …) can be
    // re-bound by later compile/link passes (ESM import wiring, namespace
    // rebuilds) to a fresh, unwired object — which breaks `x.constructor
    // === Array` and `getPrototypeOf(x) === Array.prototype` identity.
    // These `__ctor_<Name>` globals are the constructor-side companion to
    // `__tid_<Name>` (the type stamp): a stable, host-owned anchor to the
    // ONE canonical constructor (the same object wired into the shared
    // prototype's `.constructor`). The compiler resolves bare builtin
    // constructor *values* through these, so identity is immune to the
    // user-facing global being clobbered.
    // Core builtins whose prototype is a process-global singleton: read the
    // canonical (first-writer-wins) constructor straight off the shared
    // prototype, so `__ctor_<Name>` matches `x.constructor` across parallel
    // VMs even though each VM minted its own `<Name>` global.
    let core_protos: [(&str, Value); 6] = [
        ("Object", crate::ecma::object::shared_object_prototype()),
        ("Array", crate::ecma::array::shared_array_prototype()),
        (
            "Function",
            crate::ecma::function::shared_function_prototype(),
        ),
        ("Number", crate::ecma::number::shared_number_prototype()),
        ("String", crate::ecma::string::shared_string_prototype()),
        ("Boolean", crate::ecma::boolean::shared_boolean_prototype()),
    ];
    for (name, proto) in &core_protos {
        if let Value::Object(p) = proto {
            if let Some(ctor) = p.lock().unwrap().properties.get("constructor").cloned() {
                vm.globals.insert(format!("__ctor_{name}"), ctor);
            }
        }
    }
    // Remaining builtins (no shared-prototype singleton): per-VM global is fine.
    for name in &["Symbol", "BigInt", "Date", "RegExp"] {
        if let Some(ctor) = vm.globals.get(*name).cloned() {
            vm.globals.insert(format!("__ctor_{name}"), ctor);
        }
    }

    // Built-in Error hierarchy (ECMA-262 §20.5). These are the same canonical
    // constructor objects `ecma:value.constructorOf` returns for error
    // instances, so `e.constructor === TypeError` holds. The compiler resolves
    // the bare `TypeError` / `Error` / … value through these `__ctor_<Name>`
    // anchors.
    for name in &[
        "Error",
        "TypeError",
        "RangeError",
        "ReferenceError",
        "SyntaxError",
        "URIError",
        "EvalError",
        "AggregateError",
    ] {
        let ctor = crate::ecma::value::error_constructor_for(name);
        vm.globals.insert(format!("__ctor_{name}"), ctor);
    }
}

/// Build a Value that is callable as a host function AND can carry
/// static properties — like real JS where `Symbol` and `BigInt` are
/// both callable AND have static methods (`Symbol.for`, `BigInt.asIntN`).
fn host_fn_with_props(vm: &VM, module: &str, name: &str) -> Value {
    host_fn_ref(vm, module, name)
}
