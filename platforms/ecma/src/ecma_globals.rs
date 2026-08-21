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

use crate::receiver_host_fn_ref;
use std::sync::Arc;
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{VM, Value};

// ---- Namespace/global wiring helpers (moved from vybe_host::namespaces) ----

/// Ensure a namespace object exists at the given dotted path on the VM globals.
/// Creates intermediate objects as needed. Returns the leaf object. Stores the
/// root global under BOTH the original case and lowercase (case-sensitive
/// languages hit the original key, case-insensitive ones the lowercase).
fn ensure_namespace(vm: &mut VM, path: &[&str]) -> Value {
    if path.is_empty() {
        return Value::Null;
    }
    let root_orig = path[0].to_string();
    let root_lc = root_orig.to_lowercase();
    let root = if let Some(existing) = vm
        .global(&root_orig)
        .or_else(|| vm.global(&root_lc))
    {
        existing.clone()
    } else {
        let obj = Value::Object(vybe_runtime::heap::alloc(Object::new()));
        vm.set_global_owned(root_orig.clone(), obj.clone());
        if root_lc != root_orig {
            vm.set_global_owned(root_lc, obj.clone());
        }
        obj
    };
    let mut current = root;
    for &segment in &path[1..] {
        let orig = segment.to_string();
        let key_lc = orig.to_lowercase();
        let next = if let Value::Object(ref obj) = current {
            let lock = obj.lock().unwrap();
            lock.properties
                .get(&orig)
                .or_else(|| lock.properties.get(&key_lc))
                .cloned()
        } else {
            None
        };
        if let Some(existing) = next {
            current = existing;
        } else {
            let new_obj = Value::Object(vybe_runtime::heap::alloc(Object::new()));
            if let Value::Object(ref obj) = current {
                let mut o = obj.lock().unwrap();
                o.properties.insert(orig.clone(), new_obj.clone());
                if key_lc != orig {
                    o.properties.insert(key_lc, new_obj.clone());
                }
            }
            current = new_obj;
        }
    }
    current
}

/// Set a property on a namespace object, under BOTH original-case and
/// lowercase keys (same underlying Value).
///
/// NON-ENUMERABLE, because §17 makes that the DEFAULT for everything this
/// function installs: "Every other data property described in clauses
/// [Global Object] through [Reflection] … has the attributes { [[Writable]]:
/// *true*, [[Enumerable]]: *false*, [[Configurable]]: *true* } unless
/// otherwise specified" — and nothing in clauses 19–28 specifies otherwise for
/// a built-in prototype method or constructor property. Marking here rather
/// than at each call site is what makes the rule hold: enumerability used to
/// depend on every caller remembering a separate `track_nonenum`, and they
/// disagreed — Array marked both spellings, Map/Set marked only the exact one,
/// Date marked nothing. §EnumerateObjectProperties reads attributes through
/// [[GetOwnProperty]] and walks the prototype chain, so each miss surfaced as
/// mount plumbing leaking out of a `for...in`: `new Date(0)` enumerated 72
/// keys where the standard requires none.
///
/// The lowercase alias is marked too. It is not a spec property at all — it
/// exists so case-insensitive languages can resolve `getfullyear` — so it must
/// never be enumerable in ANY language.
fn set_prop(ns: &Value, name: &str, value: Value) {
    if let Value::Object(obj) = ns {
        let lc = name.to_lowercase();
        {
            let mut o = obj.lock().unwrap();
            o.properties.insert(name.to_string(), value.clone());
            if lc != name {
                o.properties.insert(lc.clone(), value);
            }
        }
        crate::object::track_nonenum(obj, name);
        if lc != name {
            crate::object::track_nonenum(obj, &lc);
        }
    }
}

/// Wire a shared prototype singleton's `constructor` exactly once
/// (first-writer-wins) and return the canonical constructor now on it. Pinning
/// on first write makes `x.constructor === Ctor` hold across parallel VMs.
fn set_constructor_once(proto: &Value, ctor: Value) -> Value {
    if let Value::Object(obj) = proto {
        {
            let mut o = obj.lock().unwrap();
            if let Some(existing) = o.properties.get("constructor") {
                return existing.clone();
            }
            o.properties.insert("constructor".to_string(), ctor.clone());
        }
        // §ClassDefinitionEvaluation installs `constructor` with
        // `DefineMethodProperty(proto, "constructor", F, *false*)` — the
        // trailing *false* IS the [[Enumerable]] attribute. Marked here so the
        // rule holds for every prototype, not only the ones whose call site
        // remembered.
        crate::object::track_nonenum(obj, "constructor");
    }
    ctor
}

/// Create a HostFunction Value referencing a registered host function. Stamps
/// the shared ecma function prototype, so it lives with ecma.
fn host_fn_ref(vm: &VM, module: &str, name: &str) -> Value {
    if let Some(&idx) = vm
        .host_registry
        .get(&(module.to_string(), name.to_string()))
    {
        let mut obj = Object::new();
        obj.properties
            .insert("__host_module".into(), Value::String(Arc::from(module)));
        obj.properties
            .insert("__host_name".into(), Value::String(Arc::from(name)));
        obj.properties
            .insert("__host_idx".into(), Value::F64(idx as f64));
        obj.properties.insert(
            "__proto__".into(),
            crate::function::shared_function_prototype(),
        );
        obj.properties
            .insert("name".into(), Value::String(Arc::from(name)));
        obj.kind = ObjectKind::HostFunction(idx);
        Value::Object(vybe_runtime::heap::alloc(obj))
    } else {
        Value::Null
    }
}

/// Install a built-in constructor's `prototype` data property with the
/// ECMA-262 attributes { [[Writable]]: false, [[Enumerable]]: false,
/// [[Configurable]]: false } (§20.1.2 and the per-constructor definitions).
///
/// The `[[Configurable]]: false` part is load-bearing for process safety, not
/// just spec fidelity: the canonical constructor pinned on a shared prototype
/// singleton (`set_constructor_once`) is reachable as `__ctor_<Name>` from
/// every VM, so a program doing `delete Object.prototype` would otherwise strip
/// `prototype` off the shared object and corrupt every subsequent VM. Marking
/// it non-configurable makes `ecma:object.delete` (which honors `__nonconfig`)
/// a no-op, matching the spec and containing the mutation.
fn set_ctor_prototype(ctor: &Value, proto: Value) {
    set_prop(ctor, "prototype", proto);
    if let Value::Object(obj) = ctor {
        crate::object::track_nonconfig(obj, "prototype");
        crate::object::track_nonenum(obj, "prototype");
    }
}

pub fn register(vm: &mut VM) {
    // ── Object / boxed primitive constructors ─────────────────────
    let object = host_fn_ref(vm, "ecma:object", "Object");
    set_prop(&object, "name", Value::String(Arc::from("Object")));
    let object_proto = crate::object::shared_object_prototype();
    set_constructor_once(&object_proto, object.clone());
    if let Value::Object(proto) = &object_proto {
        crate::object::track_nonenum(proto, "constructor");
        crate::object::track_nonenum(proto, "constructor");
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
            crate::object::track_nonenum(proto, name);
            let lower = name.to_lowercase();
            if lower != *name {
                crate::object::track_nonenum(proto, &lower);
            }
        }
    }
    set_ctor_prototype(&object, object_proto.clone());
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
        let value = host_fn_ref(vm, "ecma:object", name);
        if *name == "hasOwn" {
            set_prop(&value, "length", Value::I32(2));
        }
        set_prop(&object, name, value);
        if let Value::Object(object_obj) = &object {
            crate::object::track_nonenum(object_obj, name);
        }
    }
    set_prop(&object, "groupBy", Value::Bool(true));
    vm.set_global_owned("Object".to_string(), object.clone());
    vm.set_global_owned("object".to_string(), object.clone());

    let number = host_fn_ref(vm, "ecma:number", "Number");
    set_prop(&number, "name", Value::String(Arc::from("Number")));
    let number_proto = crate::number::shared_number_prototype();
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
    set_ctor_prototype(&number, number_proto);
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
    vm.set_global_owned("Number".to_string(), number.clone());
    vm.set_global_owned("number".to_string(), number.clone());

    let string = host_fn_ref(vm, "ecma:string", "String");
    set_prop(&string, "name", Value::String(Arc::from("String")));
    let string_proto = crate::string::shared_string_prototype();
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
    set_ctor_prototype(&string, string_proto);
    for name in &["fromCharCode", "fromCodePoint", "raw"] {
        set_prop(&string, name, host_fn_ref(vm, "ecma:string", name));
    }
    vm.set_global_owned("String".to_string(), string.clone());
    vm.set_global_owned("string".to_string(), string.clone());

    let boolean = host_fn_ref(vm, "ecma:boolean", "Boolean");
    set_prop(&boolean, "name", Value::String(Arc::from("Boolean")));
    let boolean_proto = crate::boolean::shared_boolean_prototype();
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
    set_ctor_prototype(&boolean, boolean_proto);
    vm.set_global_owned("Boolean".to_string(), boolean.clone());
    vm.set_global_owned("boolean".to_string(), boolean.clone());

    let function = Value::Object(vybe_runtime::heap::alloc(Object::new()));
    set_prop(&function, "name", Value::String(Arc::from("Function")));
    let function_proto = crate::function::shared_function_prototype();
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
    set_ctor_prototype(&function, function_proto.clone());
    set_prop(&function, "__proto__", function_proto);
    vm.set_global_owned("Function".to_string(), function.clone());
    vm.set_global_owned("function".to_string(), function.clone());

    let array = host_fn_ref(vm, "ecma:array", "new");
    set_prop(&array, "name", Value::String(Arc::from("Array")));
    set_prop(
        &array,
        "__proto__",
        crate::function::shared_function_prototype(),
    );
    let array_proto = crate::array::shared_array_prototype();
    set_constructor_once(&array_proto, array.clone());
    if let Value::Object(proto) = &array_proto {
        crate::object::track_nonenum(proto, "constructor");
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
            crate::object::track_nonenum(proto, name);
            let lower = name.to_lowercase();
            if lower != *name {
                crate::object::track_nonenum(proto, &lower);
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
            crate::object::track_nonenum(proto, "iterator");
            crate::object::track_nonenum(proto, "iterator");
        }
    }
    set_ctor_prototype(&array, array_proto);
    for name in &["from", "fromAsync", "isArray", "of"] {
        set_prop(&array, name, host_fn_ref(vm, "ecma:array", name));
    }
    vm.set_global_owned("Array".to_string(), array.clone());
    vm.set_global_owned("array".to_string(), array.clone());

    for (global_name, module, methods) in [
        (
            "Map",
            "ecma:map",
            &[
                "clear",
                "delete",
                "entries",
                "forEach",
                "get",
                "has",
                "keys",
                "set",
                "size",
                "values",
                "getOrInsert",
                "getOrInsertComputed",
            ][..],
        ),
        (
            "Set",
            "ecma:set",
            &[
                "add",
                "clear",
                "delete",
                "difference",
                "entries",
                "forEach",
                "has",
                "intersection",
                "isDisjointFrom",
                "isSubsetOf",
                "isSupersetOf",
                "keys",
                "size",
                "symmetricDifference",
                "union",
                "values",
            ][..],
        ),
    ] {
        let ctor = host_fn_ref(vm, module, "new");
        if !matches!(ctor, Value::Null) {
            set_prop(&ctor, "name", Value::String(Arc::from(global_name)));
            set_prop(
                &ctor,
                "__proto__",
                crate::function::shared_function_prototype(),
            );
            // The SHARED prototype singleton, not a fresh object per VM:
            // §24.1.4 / §24.2.4 — "Map/Set instances are ordinary objects that
            // inherit properties from %Map.prototype% / %Set.prototype%". The
            // instances are linked to that singleton when they are created, so
            // this must be the same object.
            let proto = if global_name == "Map" {
                crate::map::shared_map_prototype()
            } else {
                crate::set::shared_set_prototype()
            };
            set_prop(&proto, "__proto__", object_proto.clone());
            set_constructor_once(&proto, ctor.clone());
            if let Value::Object(ref p) = proto {
                crate::object::track_nonenum(p, "constructor");
            }
            for method in methods {
                if let Some(&idx) = vm
                    .host_registry
                    .get(&(module.to_string(), (*method).to_string()))
                {
                    // §24.1.3.10 / §24.2.3.9 — `size` is an ACCESSOR whose set
                    // accessor is undefined, not a method. `__get_<name>` is
                    // how the VM spells a getter (`dispatch.rs`), so the same
                    // host function is installed there instead of under `size`.
                    let key = if *method == "size" {
                        "__get_size"
                    } else {
                        *method
                    };
                    set_prop(&proto, key, receiver_host_fn_ref(module, method, idx));
                    if let Value::Object(ref p) = proto {
                        crate::object::track_nonenum(p, key);
                        // An accessor stored as `__get_size` is EXPOSED under
                        // its bare name, so the attribute has to be recorded
                        // against that spelling too — §24.1.3.10 makes `size`
                        // itself { [[Enumerable]]: false, [[Configurable]]:
                        // true }. Marking only the storage key left the two
                        // readers disagreeing: `propertyIsEnumerable("size")`
                        // answered false while `Object.keys` and `for...in`
                        // still yielded it.
                        if key != *method {
                            crate::object::track_nonenum(p, method);
                        }
                    }
                }
            }
            set_ctor_prototype(&ctor, proto);
            vm.set_global_owned(global_name.to_string(), ctor.clone());
            vm.set_global_owned(global_name.to_ascii_lowercase().to_string(), ctor);
        }
    }

    let date = host_fn_ref(vm, "ecma:date", "new");
    if !matches!(date, Value::Null) {
        set_prop(&date, "name", Value::String(Arc::from("Date")));
        set_prop(
            &date,
            "__proto__",
            crate::function::shared_function_prototype(),
        );
        let date_proto = crate::date::shared_date_prototype();
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
        set_ctor_prototype(&date, date_proto);
        for name in &["now", "parse", "UTC"] {
            set_prop(&date, name, host_fn_ref(vm, "ecma:date", name));
        }
        vm.set_global_owned("Date".to_string(), date.clone());
        vm.set_global_owned("date".to_string(), date.clone());
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
    vm.set_global_owned("Symbol".to_string(), sym);
    vm.set_global_owned(
        "symbol".to_string(),
        vm.global("Symbol").cloned().unwrap_or(Value::Null),
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
    vm.set_global_owned("BigInt".to_string(), bigint);
    vm.set_global_owned(
        "bigint".to_string(),
        vm.global("BigInt").cloned().unwrap_or(Value::Null),
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
        vm.set_global_owned((*name).to_string(), host_fn_ref(vm, "ecma:string", name));
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
            if let Some(&idx) = vm
                .host_registry
                .get(&(module.to_string(), "from".to_string()))
            {
                set_prop(&ctor, "from", receiver_host_fn_ref(module, "from", idx));
            }
            if let Some(&idx) = vm
                .host_registry
                .get(&(module.to_string(), "of".to_string()))
            {
                set_prop(&ctor, "of", receiver_host_fn_ref(module, "of", idx));
            }
            set_prop(&ctor, "BYTES_PER_ELEMENT", Value::I32(*bpe));
            set_prop(&ctor, "__vybe_typed_array_ctor", Value::Bool(true));
            let proto = Value::Object(vybe_runtime::heap::alloc(Object::new()));
            set_prop(&proto, "__proto__", object_proto.clone());
            for method in [
                "at",
                "copyWithin",
                "entries",
                "every",
                "fill",
                "filter",
                "find",
                "findIndex",
                "findLast",
                "findLastIndex",
                "forEach",
                "includes",
                "indexOf",
                "join",
                "keys",
                "lastIndexOf",
                "map",
                "reduce",
                "reduceRight",
                "reverse",
                "set",
                "slice",
                "some",
                "sort",
                "subarray",
                "toLocaleString",
                "toReversed",
                "toSorted",
                "toString",
                "values",
                "with",
            ] {
                if let Some(&idx) = vm
                    .host_registry
                    .get(&(module.to_string(), method.to_string()))
                {
                    set_prop(&proto, method, receiver_host_fn_ref(module, method, idx));
                }
            }
            set_constructor_once(&proto, ctor.clone());
            if let Value::Object(p) = &proto {
                crate::object::track_nonenum(p, "constructor");
            }
            set_ctor_prototype(&ctor, proto);
            vm.set_global_owned(global_name.to_string(), ctor.clone());
            vm.set_global_owned(global_name.to_ascii_lowercase().to_string(), ctor);
        }
    }

    // ── ArrayBuffer / SharedArrayBuffer / DataView constructors ─────────
    // These are primarily emitted through known_types for `new X(...)`, but
    // code also feature-probes `X.prototype.foo`. Provide an ordinary
    // prototype object so missing Stage-4 methods read as `undefined` instead
    // of throwing while accessing `.prototype`.
    for (global_name, module) in &[
        ("ArrayBuffer", "ecma:arraybuffer"),
        ("SharedArrayBuffer", "ecma:sharedarraybuffer"),
        ("DataView", "ecma:dataview"),
    ] {
        let ctor = host_fn_ref(vm, module, "new");
        if !matches!(ctor, Value::Null) {
            set_prop(&ctor, "name", Value::String(Arc::from(*global_name)));
            let proto = Value::Object(vybe_runtime::heap::alloc(Object::new()));
            set_prop(&proto, "__proto__", object_proto.clone());
            for method in ["slice", "resize", "transfer", "transferToFixedLength"] {
                if let Some(&idx) = vm
                    .host_registry
                    .get(&(module.to_string(), method.to_string()))
                {
                    set_prop(&proto, method, receiver_host_fn_ref(module, method, idx));
                }
            }
            set_constructor_once(&proto, ctor.clone());
            if let Value::Object(p) = &proto {
                crate::object::track_nonenum(p, "constructor");
            }
            set_ctor_prototype(&ctor, proto);
            vm.set_global_owned(global_name.to_string(), ctor);
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
        let regexp_proto = crate::regexp::shared_regexp_prototype();
        set_constructor_once(&regexp_proto, regexp.clone());
        if let Value::Object(p) = &regexp_proto {
            crate::object::track_nonenum(p, "constructor");
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
                crate::object::track_nonenum(p, name);
            }
        }
        set_ctor_prototype(&regexp, regexp_proto);
        vm.set_global_owned("RegExp".to_string(), regexp.clone());
        vm.set_global_owned("regexp".to_string(), regexp);
    }

    // ── globalThis — proper §19.3.1 singleton ──────────────────────
    // Pulls the shared process-global Object that `ecma:globalThis.get`
    // also returns, so identity holds across both access patterns.
    vm.set_global_owned(
        "globalThis".to_string(),
        crate::global_this::shared_singleton(),
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
        ("Object", crate::object::shared_object_prototype()),
        ("Array", crate::array::shared_array_prototype()),
        ("Function", crate::function::shared_function_prototype()),
        ("Number", crate::number::shared_number_prototype()),
        ("String", crate::string::shared_string_prototype()),
        ("Boolean", crate::boolean::shared_boolean_prototype()),
    ];
    for (name, proto) in &core_protos {
        if let Value::Object(p) = proto {
            if let Some(ctor) = p.lock().unwrap().properties.get("constructor").cloned() {
                vm.set_global_owned(format!("__ctor_{name}"), ctor);
            }
        }
    }
    // Remaining builtins (no shared-prototype singleton): per-VM global is fine.
    for name in &[
        "Symbol",
        "BigInt",
        "Date",
        "RegExp",
        "Map",
        "Set",
        "ArrayBuffer",
        "SharedArrayBuffer",
        "DataView",
        "Int8Array",
        "Uint8Array",
        "Uint8ClampedArray",
        "Int16Array",
        "Uint16Array",
        "Int32Array",
        "Uint32Array",
        "Float32Array",
        "Float64Array",
        "BigInt64Array",
        "BigUint64Array",
    ] {
        if let Some(ctor) = vm.global(*name).cloned() {
            vm.set_global_owned(format!("__ctor_{name}"), ctor);
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
        let ctor = crate::value::error_constructor_for(name);
        vm.set_global_owned(format!("__ctor_{name}"), ctor);
    }

    // ── Intl (ECMA-402) — the `Intl` global + its service constructors ──
    register_intl(vm);
}

/// Wire the `Intl` global: expose each ECMA-402 service constructor
/// (`Intl.Collator`, `Intl.NumberFormat`, …) resolving to the corresponding
/// `ecma:intl/<class>:new` host fn, stamp instance methods onto each shared
/// service prototype, and add the `Intl.getCanonicalLocales` /
/// `Intl.supportedValuesOf` statics. Ported from the retired
/// `vybe_host::namespaces::intl` so `Intl` registers the same way as the rest
/// of the ecma globals (Map/Set/TypedArrays) — see the module doc.
fn register_intl(vm: &mut VM) {
    let intl = ensure_namespace(vm, &["Intl"]);

    // Constructors — `new Intl.X(...)` resolves to `ecma:intl/x:new`.
    let collator = host_fn_ref(vm, "ecma:intl/collator", "new");
    let number_format = host_fn_ref(vm, "ecma:intl/numberformat", "new");
    let date_time_format = host_fn_ref(vm, "ecma:intl/datetimeformat", "new");
    set_prop(
        &intl,
        "ListFormat",
        host_fn_ref(vm, "ecma:intl/listformat", "new"),
    );
    set_prop(
        &intl,
        "PluralRules",
        host_fn_ref(vm, "ecma:intl/pluralrules", "new"),
    );
    let relative_time_format = host_fn_ref(vm, "ecma:intl/relativetimeformat", "new");
    let segmenter = host_fn_ref(vm, "ecma:intl/segmenter", "new");
    set_prop(&intl, "Locale", host_fn_ref(vm, "ecma:intl/locale", "new"));
    set_prop(
        &intl,
        "DisplayNames",
        host_fn_ref(vm, "ecma:intl/displaynames", "new"),
    );
    set_prop(
        &intl,
        "DurationFormat",
        host_fn_ref(vm, "ecma:intl/durationformat", "new"),
    );
    set_prop(&intl, "Collator", collator.clone());
    set_prop(&intl, "NumberFormat", number_format.clone());
    set_prop(&intl, "DateTimeFormat", date_time_format.clone());
    set_prop(&intl, "RelativeTimeFormat", relative_time_format.clone());
    set_prop(&intl, "Segmenter", segmenter.clone());

    let object_proto = crate::object::shared_object_prototype();

    let collator_proto = crate::intl::shared_collator_prototype();
    set_prop(&collator_proto, "constructor", collator.clone());
    set_prop(&collator_proto, "__proto__", object_proto.clone());
    for name in &["compare", "resolvedOptions"] {
        let idx = *vm
            .host_registry
            .get(&("ecma:intl/collator".to_string(), (*name).to_string()))
            .expect("ecma:intl/collator method must be registered");
        set_prop(
            &collator_proto,
            name,
            receiver_host_fn_ref("ecma:intl/collator", name, idx),
        );
    }
    set_prop(&collator, "prototype", collator_proto);

    let number_format_proto = crate::intl::shared_number_format_prototype();
    set_prop(&number_format_proto, "constructor", number_format.clone());
    set_prop(&number_format_proto, "__proto__", object_proto.clone());
    for name in &["format", "formatToParts", "resolvedOptions"] {
        let idx = *vm
            .host_registry
            .get(&("ecma:intl/numberformat".to_string(), (*name).to_string()))
            .expect("ecma:intl/numberformat method must be registered");
        set_prop(
            &number_format_proto,
            name,
            receiver_host_fn_ref("ecma:intl/numberformat", name, idx),
        );
    }
    set_prop(&number_format, "prototype", number_format_proto);

    let date_time_format_proto = crate::intl::shared_date_time_format_prototype();
    set_prop(
        &date_time_format_proto,
        "constructor",
        date_time_format.clone(),
    );
    set_prop(&date_time_format_proto, "__proto__", object_proto.clone());
    for name in &[
        "format",
        "formatToParts",
        "formatRange",
        "formatRangeToParts",
        "resolvedOptions",
    ] {
        let idx = *vm
            .host_registry
            .get(&("ecma:intl/datetimeformat".to_string(), (*name).to_string()))
            .expect("ecma:intl/datetimeformat method must be registered");
        set_prop(
            &date_time_format_proto,
            name,
            receiver_host_fn_ref("ecma:intl/datetimeformat", name, idx),
        );
    }
    set_prop(
        &date_time_format,
        "supportedLocalesOf",
        host_fn_ref(vm, "ecma:intl/datetimeformat", "supportedLocalesOf"),
    );
    set_prop(&date_time_format, "prototype", date_time_format_proto);

    let relative_time_format_proto = crate::intl::shared_relative_time_format_prototype();
    set_prop(
        &relative_time_format_proto,
        "constructor",
        relative_time_format.clone(),
    );
    set_prop(
        &relative_time_format_proto,
        "__proto__",
        object_proto.clone(),
    );
    for name in &["format", "formatToParts", "resolvedOptions"] {
        let idx = *vm
            .host_registry
            .get(&(
                "ecma:intl/relativetimeformat".to_string(),
                (*name).to_string(),
            ))
            .expect("ecma:intl/relativetimeformat method must be registered");
        set_prop(
            &relative_time_format_proto,
            name,
            receiver_host_fn_ref("ecma:intl/relativetimeformat", name, idx),
        );
    }
    set_prop(
        &relative_time_format,
        "prototype",
        relative_time_format_proto,
    );

    let segmenter_proto = crate::intl::shared_segmenter_prototype();
    set_prop(&segmenter_proto, "constructor", segmenter.clone());
    set_prop(&segmenter_proto, "__proto__", object_proto);
    for name in &["segment", "resolvedOptions"] {
        let idx = *vm
            .host_registry
            .get(&("ecma:intl/segmenter".to_string(), (*name).to_string()))
            .expect("ecma:intl/segmenter method must be registered");
        set_prop(
            &segmenter_proto,
            name,
            receiver_host_fn_ref("ecma:intl/segmenter", name, idx),
        );
    }
    set_prop(&segmenter, "prototype", segmenter_proto);

    // Static methods — `Intl.getCanonicalLocales(...)` → `ecma:intl:*`.
    set_prop(
        &intl,
        "getCanonicalLocales",
        host_fn_ref(vm, "ecma:intl", "getCanonicalLocales"),
    );
    set_prop(
        &intl,
        "supportedValuesOf",
        host_fn_ref(vm, "ecma:intl", "supportedValuesOf"),
    );
}

/// Build a Value that is callable as a host function AND can carry
/// static properties — like real JS where `Symbol` and `BigInt` are
/// both callable AND have static methods (`Symbol.for`, `BigInt.asIntN`).
fn host_fn_with_props(vm: &VM, module: &str, name: &str) -> Value {
    host_fn_ref(vm, module, name)
}
