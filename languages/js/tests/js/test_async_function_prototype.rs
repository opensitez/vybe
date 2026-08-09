// AsyncFunction.prototype — async function metadata, prototype chain, and instanceof relations.
crate::js_cases! {
    async_declaration_prototype_is_async_function_prototype => {
        r#"async function run() {} console.log(Object.getPrototypeOf(run) === AsyncFunction.prototype);"#,
        ["true"]
    };

    async_expression_prototype_is_async_function_prototype => {
        r#"const run = async function() {}; console.log(Object.getPrototypeOf(run) === AsyncFunction.prototype);"#,
        ["true"]
    };

    async_arrow_prototype_is_async_function_prototype => {
        r#"const run = async () => {}; console.log(Object.getPrototypeOf(run) === AsyncFunction.prototype);"#,
        ["true"]
    };

    async_method_prototype_is_async_function_prototype => {
        r#"const obj = { async m() {} }; console.log(Object.getPrototypeOf(obj.m) === AsyncFunction.prototype);"#,
        ["true"]
    };

    async_class_method_prototype_is_async_function_prototype => {
        r#"class S { async load() {} } console.log(Object.getPrototypeOf(S.prototype.load) === AsyncFunction.prototype);"#,
        ["true"]
    };

    async_static_method_prototype_is_async_function_prototype => {
        r#"class S { static async load() {} } console.log(Object.getPrototypeOf(S.load) === AsyncFunction.prototype);"#,
        ["true"]
    };

    async_function_instanceof_async_function => {
        r#"async function f() {} console.log(f instanceof AsyncFunction);"#,
        ["true"]
    };

    async_arrow_instanceof_async_function => {
        r#"const f = async () => {}; console.log(f instanceof AsyncFunction);"#,
        ["true"]
    };

    async_function_instanceof_function => {
        r#"async function f() {} console.log(f instanceof Function);"#,
        ["true"]
    };

    async_function_not_instanceof_generator_function => {
        r#"async function f() {} console.log(f instanceof GeneratorFunction);"#,
        ["false"]
    };

    async_function_not_instanceof_async_generator_function => {
        r#"async function f() {} console.log(f instanceof AsyncGeneratorFunction);"#,
        ["false"]
    };

    regular_function_not_instanceof_async_function => {
        r#"function f() {} console.log(f instanceof AsyncFunction);"#,
        ["false"]
    };

    async_function_prototype_prototype_is_function_prototype => {
        r#"console.log(Object.getPrototypeOf(AsyncFunction.prototype) === Function.prototype);"#,
        ["true"]
    };

    async_function_constructor_prototype_is_function_prototype => {
        r#"console.log(Object.getPrototypeOf(AsyncFunction) === Function.prototype);"#,
        ["true"]
    };

    async_function_constructor_name_is_async_function => {
        r#"console.log(AsyncFunction.name);"#,
        ["AsyncFunction"]
    };

    async_function_prototype_name_is_empty => {
        r#"console.log(AsyncFunction.prototype.name);"#,
        [""]
    };

    async_function_prototype_length_is_zero => {
        r#"console.log(AsyncFunction.prototype.length);"#,
        ["0"]
    };

    async_function_length_counts_simple_params => {
        r#"async function f(a, b, c) {} console.log(f.length);"#,
        ["3"]
    };

    async_function_length_stops_at_default => {
        r#"async function f(a, b = 1) {} console.log(f.length);"#,
        ["1"]
    };

    async_arrow_length_with_rest_only_is_zero => {
        r#"const f = async (...xs) => xs; console.log(f.length);"#,
        ["0"]
    };

    async_function_to_string_contains_async_keyword => {
        r#"async function ping() {} console.log(Function.prototype.toString.call(ping).includes("async"));"#,
        ["true"]
    };

    async_arrow_to_string_contains_async => {
        r#"const f = async () => 1; console.log(Function.prototype.toString.call(f).includes("async"));"#,
        ["true"]
    };

    async_function_call_returns_promise => {
        r#"async function value() { return 4; } console.log(value() instanceof Promise);"#,
        ["true"]
    };

    async_function_apply_returns_promise => {
        r#"async function value() { return 4; } console.log(value.apply(null, []) instanceof Promise);"#,
        ["true"]
    };

    async_function_bind_returns_bound_async_function => {
        r#"async function value() { return this.n; } const b = value.bind({ n: 2 }); console.log(b() instanceof Promise);"#,
        ["true"]
    };

    bound_async_function_prototype_is_async_function_prototype => {
        r#"async function value() {} const b = value.bind(null); console.log(Object.getPrototypeOf(b) === AsyncFunction.prototype);"#,
        ["true"]
    };

    bound_async_function_instanceof_async_function => {
        r#"async function value() {} const b = value.bind(null); console.log(b instanceof AsyncFunction);"#,
        ["true"]
    };

    async_function_has_no_own_prototype_property => {
        r#"async function f() {} console.log(f.hasOwnProperty("prototype"));"#,
        ["false"]
    };

    async_arrow_prototype_property_is_undefined => {
        r#"const f = async () => {}; console.log(f.prototype === undefined);"#,
        ["true"]
    };

    async_function_constructor_is_async_function => {
        r#"console.log(async function f() {}.constructor === AsyncFunction);"#,
        ["true"]
    };

    async_function_prototype_has_call_method => {
        r#"console.log(typeof AsyncFunction.prototype.call);"#,
        ["function"]
    };

    async_function_prototype_has_apply_method => {
        r#"console.log(typeof AsyncFunction.prototype.apply);"#,
        ["function"]
    };

    async_function_prototype_has_bind_method => {
        r#"console.log(typeof AsyncFunction.prototype.bind);"#,
        ["function"]
    };

    async_function_prototype_has_to_string => {
        r#"console.log(typeof AsyncFunction.prototype.toString);"#,
        ["function"]
    };

    async_function_prototype_inherits_from_function_prototype => {
        r#"console.log("call" in AsyncFunction.prototype && "apply" in AsyncFunction.prototype);"#,
        ["true"]
    };

    async_function_is_not_same_as_async_function_prototype => {
        r#"async function f() {} console.log(f === AsyncFunction.prototype);"#,
        ["false"]
    };

    async_function_prototype_is_not_async_function_instance => {
        r#"async function f() {} console.log(AsyncFunction.prototype instanceof AsyncFunction);"#,
        ["false"]
    };

    async_iife_prototype_is_async_function_prototype => {
        r#"const f = (async function() {})(); console.log(f instanceof Promise);"#,
        ["true"]
    };

    async_method_name_in_object_literal => {
        r#"const api = { async fetch() {} }; console.log(api.fetch.name);"#,
        ["fetch"]
    };

    async_named_expression_preserves_inner_name => {
        r#"const outer = async function inner() {}; console.log(outer.name);"#,
        ["inner"]
    };

    async_function_assigned_to_variable_infers_name => {
        r#"const worker = async function() {}; console.log(worker.name);"#,
        ["worker"]
    };

    // §27.7: an async function returns a FRESH promise per call, so the
    // two results can never be `===` even though the lexical `this` is
    // identical (node-verified false). Await the values to compare them.
    async_function_call_with_this_ignores_this_in_arrow => {
        r#"const f = async () => this; (async () => { console.log((await f.call({ x: 1 })) === (await f.call({ x: 2 }))); })();"#,
        ["true"]
    };

    async_function_call_with_this_on_regular_async => {
        r#"async function read() { return this.v; } console.log(read.call({ v: 9 }) instanceof Promise);"#,
        ["true"]
    };

    async_function_apply_with_args_returns_promise => {
        r#"async function add(a, b) { return a + b; } const p = add.apply(null, [2, 3]); console.log(p instanceof Promise);"#,
        ["true"]
    };

    async_function_bind_partial_preserves_async_prototype => {
        r#"async function add(a, b) { return a + b; } const plusOne = add.bind(null, 1); console.log(Object.getPrototypeOf(plusOne) === AsyncFunction.prototype);"#,
        ["true"]
    };

    async_function_bind_length_reduced => {
        r#"async function f(a, b) {} console.log(f.bind(null, 1).length);"#,
        ["1"]
    };

    async_function_instanceof_object_is_false => {
        r#"async function f() {} console.log(f instanceof Object);"#,
        ["true"]
    };

    async_function_prototype_not_equal_function_prototype => {
        r#"console.log(AsyncFunction.prototype === Function.prototype);"#,
        ["false"]
    };

    // §27.7.2: %AsyncFunction.prototype% is an ordinary object, not a
    // function (node-verified "object").
    async_function_prototype_is_object => {
        r#"console.log(typeof AsyncFunction.prototype);"#,
        ["object"]
    };

    async_function_prototype_constructor_is_async_function => {
        r#"console.log(AsyncFunction.prototype.constructor === AsyncFunction);"#,
        ["true"]
    };

    async_function_from_new_function_is_not_async_function => {
        r#"const f = new Function("return 1;"); console.log(f instanceof AsyncFunction);"#,
        ["false"]
    };

    generator_function_not_instanceof_async_function => {
        r#"function* g() {} console.log(g instanceof AsyncFunction);"#,
        ["false"]
    };

    async_generator_not_instanceof_async_function => {
        r#"async function* g() {} console.log(g instanceof AsyncGeneratorFunction);"#,
        ["true"]
    };

    async_function_prototype_to_string_is_function => {
        r#"console.log(typeof AsyncFunction.prototype.toString.call(async function() {}));"#,
        ["string"]
    };

    async_function_has_configurable_name => {
        r#"async function f() {} console.log(Object.getOwnPropertyDescriptor(f, "name").configurable);"#,
        ["true"]
    };

    async_function_has_configurable_length => {
        r#"async function f(a) {} console.log(Object.getOwnPropertyDescriptor(f, "length").configurable);"#,
        ["true"]
    };

    async_function_symbol_has_instance_via_function => {
        r#"async function F() {} console.log(F[Symbol.hasInstance] === Function[Symbol.hasInstance]);"#,
        ["true"]
    };

    async_arrow_assigned_to_property_infers_name => {
        r#"const obj = { task: async () => {} }; console.log(obj.task.name);"#,
        ["task"]
    };

    async_function_prototype_call_invokes_and_returns_promise => {
        r#"async function f() { return 1; } console.log(AsyncFunction.prototype.call.call(f, null) instanceof Promise);"#,
        ["true"]
    };

    async_function_prototype_apply_invokes_and_returns_promise => {
        r#"async function f() { return 1; } console.log(AsyncFunction.prototype.apply.call(f, null, []) instanceof Promise);"#,
        ["true"]
    };

    async_function_prototype_bind_creates_async_callable => {
        r#"async function f() { return 1; } const b = AsyncFunction.prototype.bind.call(f, null); console.log(b instanceof AsyncFunction);"#,
        ["true"]
    };

    object_get_prototype_of_async_function_is_not_function_prototype => {
        r#"async function f() {} console.log(Object.getPrototypeOf(f) === Function.prototype);"#,
        ["false"]
    };

    async_function_extends_function_prototype_chain => {
        r#"async function f() {} console.log(Function.prototype.isPrototypeOf(f));"#,
        ["true"]
    };

    async_function_prototype_is_prototype_of_async_instances => {
        r#"async function f() {} console.log(AsyncFunction.prototype.isPrototypeOf(f));"#,
        ["true"]
    };

    // §27.7.1: %AsyncFunction.prototype%'s [[Prototype]] is
    // %Function.prototype%, and isPrototypeOf walks the whole chain —
    // node-verified true.
    function_prototype_is_not_prototype_of_async_function_directly => {
        r#"async function f() {} console.log(Function.prototype.isPrototypeOf(f));"#,
        ["true"]
    };

    async_function_prototype_tostringtag => {
        r#"console.log(AsyncFunction.prototype[Symbol.toStringTag]);"#,
        ["AsyncFunction"]
    };
}
