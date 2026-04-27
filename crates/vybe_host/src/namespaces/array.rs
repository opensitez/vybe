use super::*;

pub fn register(vm: &mut VM) {
    // VB / .NET-shape namespace property lookups bound to the canonical
    // ECMA-262 imports. `Array.IsArray(x)`, `System.Array.Reverse(arr)`,
    // `System.Collections.Generic.Dictionary` etc. all resolve to the
    // same host fns JS / V8 satisfy via the wasm-js-builtins proposal.
    let array = ensure_namespace(vm, &["Array"]);
    set_prop(&array, "isarray", host_fn_ref(vm, "ecma:array", "isArray"));
    set_prop(&array, "from", host_fn_ref(vm, "ecma:array", "from"));

    // System.Array
    let sys = ensure_namespace(vm, &["System", "Array"]);
    set_prop(&sys, "reverse", host_fn_ref(vm, "ecma:array", "reverse"));
    set_prop(&sys, "indexof", host_fn_ref(vm, "ecma:array", "indexOf"));

    // System.Collections.Generic — `Dictionary<K,V>` is shape-identical to
    // a JS Map (IndexMap-backed; SameValueZero key equality), so the .NET
    // generic constructor binds to `ecma:map.new`. `List<T>` constructs
    // from an iterable, mirroring `Array.from`.
    let coll = ensure_namespace(vm, &["System", "Collections", "Generic"]);
    set_prop(&coll, "list", host_fn_ref(vm, "ecma:array", "from"));
    set_prop(&coll, "dictionary", host_fn_ref(vm, "ecma:map", "new"));
}
