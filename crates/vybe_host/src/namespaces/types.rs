use super::*;

pub fn register(vm: &mut VM) {
    register_datetime_ns(vm);
    register_stringbuilder_ns(vm);
    register_collections_ns(vm);
    register_timespan_ns(vm);
    register_guid_ns(vm);
    register_primitives_ns(vm);
    register_process_ns(vm);
    register_array_statics_ns(vm);
}

fn register_datetime_ns(vm: &mut VM) {
    // DateTime / System.DateTime — `Now` / `UtcNow` / `Today` /
    // `Parse` static-method dispatch handled at compile time through
    // the dotnet wrapper Component Model adapter
    // (`emitter::dotnet::core::datetime_adapter` over `ecma:date.*` /
    // `wasi:clocks/wall-clock.now`). The namespace objects are still
    // ensured so identifier resolution doesn't fail; `MaxValue` /
    // `MinValue` constants stay as plain values.
    ensure_namespace(vm, &["DateTime"]);
    let sys_dt = ensure_namespace(vm, &["System", "DateTime"]);
    set_prop(&sys_dt, "maxvalue", Value::F64(253402300799.0)); // 9999-12-31
    set_prop(&sys_dt, "minvalue", Value::F64(0.0));
}

fn register_stringbuilder_ns(vm: &mut VM) {
    // `System.Text.StringBuilder` namespace ensured (so the
    // identifier resolves) but the `.new` binding retired —
    // `new StringBuilder()` lowers at compile time through the
    // dotnet wrapper Component Model adapter
    // (`emitter::dotnet::core::stringbuilder_adapter`). No host fn
    // for the construction path.
    ensure_namespace(vm, &["System", "Text", "StringBuilder"]);
}

fn register_collections_ns(vm: &mut VM) {
    // System.Collections.Generic — Queue / Stack / HashSet / List /
    // Dictionary constructors all dispatch at compile time through
    // the dotnet wrapper Component Model. The namespace bindings +
    // bare globals retired here pointed at `vybe:types/*New` host
    // fns that are no longer needed (the dotnet adapter routes to
    // `collections.new` / `ecma:map.new` / `ecma:set.new` directly
    // and stamps `__type` via `__stamp_type`). Namespace objects are
    // still ensured so identifier resolution doesn't fail.
    ensure_namespace(vm, &["System", "Collections", "Generic"]);
    ensure_namespace(vm, &["System", "Collections", "Generic", "Queue"]);
    ensure_namespace(vm, &["System", "Collections", "Generic", "Stack"]);
    ensure_namespace(vm, &["System", "Collections", "Generic", "HashSet"]);
}

fn register_timespan_ns(vm: &mut VM) {
    // TimeSpan / System.TimeSpan factory statics retired — all
    // dispatch handled at compile time through
    // `emitter::dotnet::core::timespan_adapter` (pure inline
    // bytecode: unit-to-ms multiply + struct build). The namespace
    // objects are still ensured so identifier resolution doesn't
    // fail.
    ensure_namespace(vm, &["TimeSpan"]);
    ensure_namespace(vm, &["System", "TimeSpan"]);
}

fn register_guid_ns(vm: &mut VM) {
    // .NET `Guid` maps to UUID v4 (RFC 4122). The spec primitive is
    // `wasi:random/random.uuid()` — WASI-aligned and the closest
    // upstream-spec match per the user's preference order
    // (wasi:* > ecma:* > web:*). `Guid.Parse(s)` is a string
    // passthrough — UUIDs are already strings in this representation.
    let empty_guid = Value::String(Arc::from("00000000-0000-0000-0000-000000000000"));
    let parse_passthrough = host_fn_ref(vm, "ecma:string", "String");
    let new_uuid = host_fn_ref(vm, "wasi:random/random", "uuid");

    let guid = ensure_namespace(vm, &["Guid"]);
    set_prop(&guid, "newguid", new_uuid.clone());
    set_prop(&guid, "empty", empty_guid.clone());
    set_prop(&guid, "parse", parse_passthrough.clone());

    let sys_guid = ensure_namespace(vm, &["System", "Guid"]);
    set_prop(&sys_guid, "newguid", new_uuid);
    set_prop(&sys_guid, "empty", empty_guid);
    set_prop(&sys_guid, "parse", parse_passthrough);
}

fn register_primitives_ns(vm: &mut VM) {
    // Primitive type namespaces — VB/C# `int.Parse`, `string.Format`, etc.
    // Bound to ECMA-262 host fns (`ecma:number.parseInt`, `ecma:string.*`).
    // Language adapters in the emitter handle .NET-shape divergence
    // (e.g. `string.IsNullOrEmpty(s)` ≡ `s == null || s.length === 0` —
    // not a single ECMA call, but expressible as a stdlib adapter chunk
    // in the language layer).
    let int_ns = ensure_namespace(vm, &["int"]);
    set_prop(&int_ns, "parse", host_fn_ref(vm, "ecma:number", "parseInt"));
    set_prop(&int_ns, "tryparse", host_fn_ref(vm, "ecma:number", "parseInt"));
    set_prop(&int_ns, "maxvalue", Value::F64(i32::MAX as f64));
    set_prop(&int_ns, "minvalue", Value::F64(i32::MIN as f64));

    // `string.Format` / `string.IsNullOrEmpty` have no single ECMA call
    // target — the bindings are dropped here. Language adapters compile
    // them to inline expressions / stdlib polyfills.
    let string_ns = ensure_namespace(vm, &["string"]);
    set_prop(&string_ns, "join", host_fn_ref(vm, "ecma:array", "join"));

    let double_ns = ensure_namespace(vm, &["double"]);
    set_prop(&double_ns, "parse", host_fn_ref(vm, "ecma:number", "parseFloat"));
    set_prop(&double_ns, "nan", Value::F64(f64::NAN));
    set_prop(&double_ns, "positiveinfinity", Value::F64(f64::INFINITY));

    let bool_ns = ensure_namespace(vm, &["bool"]);
    set_prop(&bool_ns, "parse", host_fn_ref(vm, "ecma:boolean", "Boolean"));

    // System.Double
    let dbl = ensure_namespace(vm, &["System", "Double"]);
    set_prop(&dbl, "parse", host_fn_ref(vm, "ecma:number", "parseFloat"));
    set_prop(&dbl, "tryparse", host_fn_ref(vm, "ecma:number", "parseFloat"));
    set_prop(&dbl, "maxvalue", Value::F64(f64::MAX));
    set_prop(&dbl, "minvalue", Value::F64(f64::MIN));
    set_prop(&dbl, "nan", Value::F64(f64::NAN));
    set_prop(&dbl, "positiveinfinity", Value::F64(f64::INFINITY));
    set_prop(&dbl, "negativeinfinity", Value::F64(f64::NEG_INFINITY));

    // System.Single
    let sng = ensure_namespace(vm, &["System", "Single"]);
    set_prop(&sng, "parse", host_fn_ref(vm, "ecma:number", "parseFloat"));
    set_prop(&sng, "maxvalue", Value::F64(f32::MAX as f64));
    set_prop(&sng, "minvalue", Value::F64(f32::MIN as f64));

    // System.Boolean
    let bln = ensure_namespace(vm, &["System", "Boolean"]);
    set_prop(&bln, "parse", host_fn_ref(vm, "ecma:boolean", "Boolean"));

    // System.Decimal (no ECMA decimal type — alias to Number)
    let dec = ensure_namespace(vm, &["System", "Decimal"]);
    set_prop(&dec, "parse", host_fn_ref(vm, "ecma:number", "parseFloat"));

    // System.DBNull
    let dbnull = ensure_namespace(vm, &["System", "DBNull"]);
    set_prop(&dbnull, "value", Value::Null);

    // System.EventArgs
    let ea = ensure_namespace(vm, &["System", "EventArgs"]);
    set_prop(&ea, "empty", Value::Null);
}

fn register_process_ns(vm: &mut VM) {
    // Process / System.Diagnostics.Process namespace `start` bindings
    // retired — handled at compile time by
    // `emitter::dotnet::core::process_adapter`. Namespace objects are
    // still ensured so identifier resolution doesn't fail.
    ensure_namespace(vm, &["Process"]);
    ensure_namespace(vm, &["System", "Diagnostics", "Process"]);
}

fn register_array_statics_ns(vm: &mut VM) {
    // System.Array static methods retired from this namespace
    // setup — `Array.Clear/Copy/Resize/Sort/Reverse/IndexOf` lower
    // at compile time through the dotnet wrapper Component Model
    // adapter (`emitter::dotnet::core::array_adapter`) to bundled
    // stdlib chunks that compose `ecma:array.*` primitives. The
    // namespace object itself is still ensured so identifier
    // resolution doesn't fail.
    ensure_namespace(vm, &["System", "Array"]);

    // System.Tuple
    let tuple = ensure_namespace(vm, &["System", "Tuple"]);
    set_prop(&tuple, "create", host_fn_ref(vm, "ecma:array", "from")); // simplified

    // System.BitConverter
    let bc = ensure_namespace(vm, &["System", "BitConverter"]);
    set_prop(&bc, "tostring", host_fn_ref(vm, "ecma:string", "String"));
    set_prop(&bc, "todouble", host_fn_ref(vm, "ecma:number", "Number"));
}
