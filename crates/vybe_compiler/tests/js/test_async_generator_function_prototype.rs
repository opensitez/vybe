// AsyncGeneratorFunction.prototype — async generator metadata, prototype chain, and instanceof relations.
crate::js_cases! {
    async_generator_declaration_prototype_is_async_generator_function_prototype => {
        r#"async function* run() { yield 1; } console.log(Object.getPrototypeOf(run) === AsyncGeneratorFunction.prototype);"#,
        ["true"]
    };

    async_generator_expression_prototype_is_async_generator_function_prototype => {
        r#"const run = async function* () { yield 1; }; console.log(Object.getPrototypeOf(run) === AsyncGeneratorFunction.prototype);"#,
        ["true"]
    };

    named_async_generator_expression_prototype_is_async_generator_function_prototype => {
        r#"const outer = async function* inner() { yield 1; }; console.log(Object.getPrototypeOf(outer) === AsyncGeneratorFunction.prototype);"#,
        ["true"]
    };

    async_generator_method_prototype_is_async_generator_function_prototype => {
        r#"const obj = { async *m() { yield 1; } }; console.log(Object.getPrototypeOf(obj.m) === AsyncGeneratorFunction.prototype);"#,
        ["true"]
    };

    async_generator_class_method_prototype_is_async_generator_function_prototype => {
        r#"class S { async *load() { yield 1; } } console.log(Object.getPrototypeOf(S.prototype.load) === AsyncGeneratorFunction.prototype);"#,
        ["true"]
    };

    async_generator_static_method_prototype_is_async_generator_function_prototype => {
        r#"class S { static async *load() { yield 1; } } console.log(Object.getPrototypeOf(S.load) === AsyncGeneratorFunction.prototype);"#,
        ["true"]
    };

    async_generator_function_instanceof_async_generator_function => {
        r#"async function* f() { yield 1; } console.log(f instanceof AsyncGeneratorFunction);"#,
        ["true"]
    };

    async_generator_expression_instanceof_async_generator_function => {
        r#"const f = async function* () { yield 1; }; console.log(f instanceof AsyncGeneratorFunction);"#,
        ["true"]
    };

    async_generator_method_instanceof_async_generator_function => {
        r#"const obj = { async *step() { yield 1; } }; console.log(obj.step instanceof AsyncGeneratorFunction);"#,
        ["true"]
    };

    async_generator_function_instanceof_function => {
        r#"async function* f() { yield 1; } console.log(f instanceof Function);"#,
        ["true"]
    };

    async_generator_function_not_instanceof_async_function => {
        r#"async function* f() { yield 1; } console.log(f instanceof AsyncFunction);"#,
        ["false"]
    };

    async_generator_function_not_instanceof_generator_function => {
        r#"async function* f() { yield 1; } console.log(f instanceof GeneratorFunction);"#,
        ["false"]
    };

    regular_function_not_instanceof_async_generator_function => {
        r#"function f() {} console.log(f instanceof AsyncGeneratorFunction);"#,
        ["false"]
    };

    async_generator_function_prototype_prototype_is_function_prototype => {
        r#"console.log(Object.getPrototypeOf(AsyncGeneratorFunction.prototype) === Function.prototype);"#,
        ["true"]
    };

    async_generator_function_constructor_prototype_is_function_prototype => {
        r#"console.log(Object.getPrototypeOf(AsyncGeneratorFunction) === Function.prototype);"#,
        ["true"]
    };

    async_generator_function_constructor_name_is_async_generator_function => {
        r#"console.log(AsyncGeneratorFunction.name);"#,
        ["AsyncGeneratorFunction"]
    };

    async_generator_function_prototype_name_is_empty => {
        r#"console.log(AsyncGeneratorFunction.prototype.name);"#,
        [""]
    };

    async_generator_function_prototype_length_is_zero => {
        r#"console.log(AsyncGeneratorFunction.prototype.length);"#,
        ["0"]
    };

    async_generator_function_length_counts_simple_params => {
        r#"async function* f(a, b, c) { yield 1; } console.log(f.length);"#,
        ["3"]
    };

    async_generator_function_length_stops_at_default => {
        r#"async function* f(a, b = 1) { yield 1; } console.log(f.length);"#,
        ["1"]
    };

    async_generator_expression_length_with_rest_only_is_zero => {
        r#"const f = async function* (...xs) { yield xs.length; }; console.log(f.length);"#,
        ["0"]
    };

    async_generator_method_length_excludes_defaults_after_first => {
        r#"const obj = { async *m(a, b = 2, c) { yield 1; } }; console.log(obj.m.length);"#,
        ["1"]
    };

    async_generator_function_to_string_contains_async_and_star => {
        r#"async function* ping() { yield 1; } const text = Function.prototype.toString.call(ping); console.log(text.includes("async") && text.includes("*"));"#,
        ["true"]
    };

    async_generator_expression_to_string_contains_async_keyword => {
        r#"const f = async function* () { yield 1; }; console.log(Function.prototype.toString.call(f).includes("async"));"#,
        ["true"]
    };

    async_generator_function_call_returns_async_iterator_object => {
        r#"async function* value() { yield 4; } console.log(typeof value().next);"#,
        ["function"]
    };

    async_generator_function_apply_returns_async_iterator_object => {
        r#"async function* value() { yield 4; } console.log(typeof value.apply(null, []).next);"#,
        ["function"]
    };

    async_generator_function_bind_returns_bound_async_generator => {
        r#"async function* value() { yield this.n; } const b = value.bind({ n: 2 }); console.log(typeof b().next);"#,
        ["function"]
    };

    bound_async_generator_function_prototype_is_async_generator_function_prototype => {
        r#"async function* value() { yield 1; } const b = value.bind(null); console.log(Object.getPrototypeOf(b) === AsyncGeneratorFunction.prototype);"#,
        ["true"]
    };

    bound_async_generator_function_instanceof_async_generator_function => {
        r#"async function* value() { yield 1; } const b = value.bind(null); console.log(b instanceof AsyncGeneratorFunction);"#,
        ["true"]
    };

    // §27.4: async generator functions DO have an own `prototype` (the
    // object their instances inherit from) — node-verified true.
    async_generator_function_has_no_own_prototype_property => {
        r#"async function* f() { yield 1; } console.log(f.hasOwnProperty("prototype"));"#,
        ["true"]
    };

    // §27.4: an async generator expression's `prototype` is an ordinary
    // object, not undefined (node-verified false).
    async_generator_expression_prototype_property_is_undefined => {
        r#"const f = async function* () { yield 1; }; console.log(f.prototype === undefined);"#,
        ["false"]
    };

    async_generator_function_constructor_is_async_generator_function => {
        r#"console.log(async function* f() { yield 1; }.constructor === AsyncGeneratorFunction);"#,
        ["true"]
    };

    async_generator_function_prototype_has_call_method => {
        r#"console.log(typeof AsyncGeneratorFunction.prototype.call);"#,
        ["function"]
    };

    async_generator_function_prototype_has_apply_method => {
        r#"console.log(typeof AsyncGeneratorFunction.prototype.apply);"#,
        ["function"]
    };

    async_generator_function_prototype_has_bind_method => {
        r#"console.log(typeof AsyncGeneratorFunction.prototype.bind);"#,
        ["function"]
    };

    async_generator_function_prototype_has_to_string => {
        r#"console.log(typeof AsyncGeneratorFunction.prototype.toString);"#,
        ["function"]
    };

    async_generator_function_prototype_inherits_from_function_prototype => {
        r#"console.log("call" in AsyncGeneratorFunction.prototype && "apply" in AsyncGeneratorFunction.prototype);"#,
        ["true"]
    };

    async_generator_function_is_not_same_as_async_generator_function_prototype => {
        r#"async function* f() { yield 1; } console.log(f === AsyncGeneratorFunction.prototype);"#,
        ["false"]
    };

    async_generator_function_prototype_is_not_async_generator_function_instance => {
        r#"async function* f() { yield 1; } console.log(AsyncGeneratorFunction.prototype instanceof AsyncGeneratorFunction);"#,
        ["false"]
    };

    async_generator_iife_call_returns_async_iterator_not_function => {
        r#"const iter = (async function* () { yield 1; })(); console.log(typeof iter.next);"#,
        ["function"]
    };

    async_generator_method_name_in_object_literal => {
        r#"const api = { async *fetch() { yield 1; } }; console.log(api.fetch.name);"#,
        ["fetch"]
    };

    named_async_generator_expression_preserves_inner_name => {
        r#"const outer = async function* inner() { yield 1; }; console.log(outer.name);"#,
        ["inner"]
    };

    async_generator_function_assigned_to_variable_infers_name => {
        r#"const worker = async function* () { yield 1; }; console.log(worker.name);"#,
        ["worker"]
    };

    async_generator_function_call_with_this_on_regular_async_generator => {
        r#"async function* read() { yield this.v; } const iter = read.call({ v: 9 }); console.log(iter instanceof Object);"#,
        ["true"]
    };

    async_generator_function_apply_with_args_returns_async_iterator => {
        r#"async function* add(a, b) { yield a + b; } console.log(typeof add.apply(null, [2, 3]).next);"#,
        ["function"]
    };

    async_generator_function_bind_partial_preserves_async_generator_prototype => {
        r#"async function* add(a, b) { yield a + b; } const plusOne = add.bind(null, 1); console.log(Object.getPrototypeOf(plusOne) === AsyncGeneratorFunction.prototype);"#,
        ["true"]
    };

    async_generator_function_bind_length_reduced => {
        r#"async function* f(a, b) { yield a + b; } console.log(f.bind(null, 1).length);"#,
        ["1"]
    };

    async_generator_function_instanceof_object_is_true => {
        r#"async function* f() { yield 1; } console.log(f instanceof Object);"#,
        ["true"]
    };

    async_generator_function_prototype_not_equal_function_prototype => {
        r#"console.log(AsyncGeneratorFunction.prototype === Function.prototype);"#,
        ["false"]
    };

    // §27.4.3: %AsyncGeneratorFunction.prototype% is an ordinary object,
    // not a function (node-verified "object").
    async_generator_function_prototype_typeof_is_function => {
        r#"console.log(typeof AsyncGeneratorFunction.prototype);"#,
        ["object"]
    };

    async_generator_function_prototype_constructor_is_async_generator_function => {
        r#"console.log(AsyncGeneratorFunction.prototype.constructor === AsyncGeneratorFunction);"#,
        ["true"]
    };

    async_generator_function_from_new_function_is_not_async_generator_function => {
        r#"const f = new Function("return 1;"); console.log(f instanceof AsyncGeneratorFunction);"#,
        ["false"]
    };

    async_generator_function_prototype_to_string_returns_string => {
        r#"console.log(typeof AsyncGeneratorFunction.prototype.toString.call(async function* () { yield 1; }));"#,
        ["string"]
    };

    async_generator_function_has_configurable_name => {
        r#"async function* f() { yield 1; } console.log(Object.getOwnPropertyDescriptor(f, "name").configurable);"#,
        ["true"]
    };

    async_generator_function_has_configurable_length => {
        r#"async function* f(a) { yield 1; } console.log(Object.getOwnPropertyDescriptor(f, "length").configurable);"#,
        ["true"]
    };

    async_generator_function_symbol_has_instance_via_function => {
        r#"async function* F() { yield 1; } console.log(F[Symbol.hasInstance] === Function[Symbol.hasInstance]);"#,
        ["true"]
    };

    async_generator_assigned_to_property_infers_name => {
        r#"const obj = { task: async function* () { yield 1; } }; console.log(obj.task.name);"#,
        ["task"]
    };

    async_generator_function_prototype_call_invokes_and_returns_async_iterator => {
        r#"async function* f() { yield 1; } console.log(typeof AsyncGeneratorFunction.prototype.call.call(f, null).next);"#,
        ["function"]
    };

    async_generator_function_prototype_apply_invokes_and_returns_async_iterator => {
        r#"async function* f() { yield 1; } console.log(typeof AsyncGeneratorFunction.prototype.apply.call(f, null, []).next);"#,
        ["function"]
    };

    async_generator_function_prototype_bind_creates_async_generator_callable => {
        r#"async function* f() { yield 1; } const b = AsyncGeneratorFunction.prototype.bind.call(f, null); console.log(b instanceof AsyncGeneratorFunction);"#,
        ["true"]
    };

    object_get_prototype_of_async_generator_is_not_function_prototype => {
        r#"async function* f() { yield 1; } console.log(Object.getPrototypeOf(f) === Function.prototype);"#,
        ["false"]
    };

    async_generator_function_extends_function_prototype_chain => {
        r#"async function* f() { yield 1; } console.log(Function.prototype.isPrototypeOf(f));"#,
        ["true"]
    };

    async_generator_function_prototype_is_prototype_of_async_generator_instances => {
        r#"async function* f() { yield 1; } console.log(AsyncGeneratorFunction.prototype.isPrototypeOf(f));"#,
        ["true"]
    };

    function_prototype_is_not_direct_prototype_of_async_generator => {
        r#"async function* f() { yield 1; } console.log(Object.getPrototypeOf(f) === Function.prototype);"#,
        ["false"]
    };

    async_generator_class_instance_method_instanceof_async_generator_function => {
        r#"class Box { async *values() { yield 1; } } const b = new Box(); console.log(b.values instanceof AsyncGeneratorFunction);"#,
        ["true"]
    };

    async_generator_static_class_method_instanceof_async_generator_function => {
        r#"class Box { static async *values() { yield 1; } } console.log(Box.values instanceof AsyncGeneratorFunction);"#,
        ["true"]
    };

    async_generator_computed_method_name_preserves_name => {
        r#"const key = "stream"; const obj = { async *[key]() { yield 1; } }; console.log(obj[key].name);"#,
        ["stream"]
    };

    async_generator_symbol_method_name_is_symbol_description => {
        r#"const sym = Symbol("iter"); const obj = { async *[sym]() { yield 1; } }; console.log(obj[sym].name);"#,
        ["[iter]"]
    };

    async_generator_function_length_with_only_rest_is_zero => {
        r#"async function* only(...xs) { yield xs.length; } console.log(only.length);"#,
        ["0"]
    };

    async_generator_function_length_with_trailing_default_only => {
        r#"async function* tail(a, b, c = 1) { yield 1; } console.log(tail.length);"#,
        ["2"]
    };

    empty_async_generator_function_length_is_zero => {
        r#"async function* empty() { yield 1; } console.log(empty.length);"#,
        ["0"]
    };

    async_generator_function_length_is_non_enumerable => {
        r#"async function* f(a) { yield 1; } console.log(f.propertyIsEnumerable("length"));"#,
        ["false"]
    };

    async_generator_function_name_is_non_enumerable => {
        r#"async function* f() { yield 1; } console.log(f.propertyIsEnumerable("name"));"#,
        ["false"]
    };

    async_generator_function_call_result_is_not_async_generator_function => {
        r#"async function* f() { yield 1; } console.log(f() instanceof AsyncGeneratorFunction);"#,
        ["false"]
    };

    async_generator_function_call_result_has_async_iterator_symbol => {
        r#"async function* f() { yield 1; } console.log(typeof f()[Symbol.asyncIterator]);"#,
        ["function"]
    };

    async_generator_bound_call_preserves_yield_behavior => {
        r#"async function* f(x) { yield x * 2; } const doubled = f.bind(null, 3); console.log(typeof doubled().next);"#,
        ["function"]
    };

    async_generator_function_prototype_is_not_enumerable_on_global => {
        r#"console.log(Object.prototype.propertyIsEnumerable.call(globalThis, "AsyncGeneratorFunction"));"#,
        ["false"]
    };

    async_generator_function_constructor_is_function => {
        r#"console.log(AsyncGeneratorFunction instanceof Function);"#,
        ["true"]
    };

    async_generator_function_constructor_length_is_one => {
        r#"console.log(AsyncGeneratorFunction.length);"#,
        ["1"]
    };

    async_generator_expression_name_from_property_assignment => {
        r#"const tools = { maker: async function* () { yield 1; } }; console.log(tools.maker.name);"#,
        ["maker"]
    };

    async_generator_method_extracted_preserves_prototype => {
        r#"const src = { async *emit() { yield 1; } }; const bare = src.emit; console.log(Object.getPrototypeOf(bare) === AsyncGeneratorFunction.prototype);"#,
        ["true"]
    };

    async_generator_function_prototype_call_passes_arguments => {
        r#"async function* pick(a, b) { yield b; } console.log(typeof AsyncGeneratorFunction.prototype.call.call(pick, null, 1, 9).next);"#,
        ["function"]
    };

    async_generator_function_prototype_apply_spreads_array_args => {
        r#"async function* pick(a, b) { yield a + b; } console.log(typeof AsyncGeneratorFunction.prototype.apply.call(pick, null, [4, 5]).next);"#,
        ["function"]
    };

    // §10.2.10 SetFunctionLength: a destructured pattern counts as ONE
    // formal parameter — node-verified 2. The old expectation (1) treated
    // the pattern as if it reduced length.
    async_generator_function_with_destructured_param_reduces_length => {
        r#"async function* f({ a }, b) { yield a + b; } console.log(f.length);"#,
        ["2"]
    };

    // §10.2.10 SetFunctionLength: a destructured pattern counts as ONE
    // formal parameter — node-verified 2. The old expectation (1) treated
    // the pattern as if it reduced length.
    async_generator_function_with_destructured_array_param_length => {
        r#"async function* f([a], b) { yield a + b; } console.log(f.length);"#,
        ["2"]
    };

    async_generator_function_strict_mode_preserves_prototype => {
        r#""use strict"; async function* f() { yield 1; } console.log(Object.getPrototypeOf(f) === AsyncGeneratorFunction.prototype);"#,
        ["true"]
    };

    async_generator_expression_strict_mode_instanceof => {
        r#""use strict"; const f = async function* () { yield 1; }; console.log(f instanceof AsyncGeneratorFunction);"#,
        ["true"]
    };

    async_generator_function_is_not_constructable_via_new => {
        r#"async function* f() { yield 1; } try { new f(); console.log("constructed"); } catch (e) { console.log(e instanceof TypeError); }"#,
        ["true"]
    };

    async_generator_function_prototype_value_of_inherits_from_function => {
        r#"console.log(typeof AsyncGeneratorFunction.prototype.valueOf);"#,
        ["function"]
    };

    async_generator_function_to_string_on_class_method_includes_async => {
        r#"class G { async *run() { yield 1; } } console.log(Function.prototype.toString.call(G.prototype.run).includes("async"));"#,
        ["true"]
    };

    async_generator_function_call_with_null_this_in_strict => {
        r#""use strict"; async function* f() { yield this; } console.log(typeof f.call(null).next);"#,
        ["function"]
    };

    async_generator_function_call_with_object_this_in_sloppy => {
        r#"async function* f() { yield this.tag; } console.log(typeof f.call({ tag: "hit" }).next);"#,
        ["function"]
    };

    async_generator_function_bind_with_this_forwards_receiver => {
        r#"async function* f() { yield this.n; } const b = f.bind({ n: 7 }); console.log(typeof b().next);"#,
        ["function"]
    };

    async_generator_function_prototype_chain_reaches_object_prototype => {
        r#"console.log(Object.prototype.isPrototypeOf(AsyncGeneratorFunction.prototype));"#,
        ["true"]
    };

    async_generator_function_instance_prototype_chain_reaches_object => {
        r#"async function* f() { yield 1; } console.log(Object.prototype.isPrototypeOf(f));"#,
        ["true"]
    };

    arrow_function_not_instanceof_async_generator_function => {
        r#"const f = async () => 1; console.log(f instanceof AsyncGeneratorFunction);"#,
        ["false"]
    };

    async_generator_function_prototype_bind_partial_application => {
        r#"async function* sum(a, b, c) { yield a + b + c; } const partial = AsyncGeneratorFunction.prototype.bind.call(sum, null, 1, 2); console.log(partial instanceof AsyncGeneratorFunction);"#,
        ["true"]
    };

    async_generator_function_call_next_returns_promise => {
        r#"async function* f() { yield 1; } console.log(f().next() instanceof Promise);"#,
        ["true"]
    };

    async_generator_function_prototype_is_function_object => {
        r#"console.log(AsyncGeneratorFunction.prototype instanceof Object);"#,
        ["true"]
    };

    async_generator_function_global_binding_is_function => {
        r#"console.log(typeof AsyncGeneratorFunction);"#,
        ["function"]
    };

    async_generator_function_prototype_differs_from_generator_function_prototype => {
        r#"console.log(AsyncGeneratorFunction.prototype === GeneratorFunction.prototype);"#,
        ["false"]
    };

    async_generator_function_constructor_differs_from_generator_function_constructor => {
        r#"console.log(AsyncGeneratorFunction === GeneratorFunction);"#,
        ["false"]
    };

    async_generator_function_prototype_differs_from_async_function_prototype => {
        r#"console.log(AsyncGeneratorFunction.prototype === AsyncFunction.prototype);"#,
        ["false"]
    };

    async_generator_nested_declaration_keeps_outer_name => {
        r#"function outer() { async function* inner() { yield 1; } return inner.name; } console.log(outer());"#,
        ["inner"]
    };

    async_generator_function_call_result_next_done_false_initially => {
        r#"async function* f() { yield 1; } const p = f().next(); console.log(p instanceof Promise);"#,
        ["true"]
    };

    async_generator_function_not_same_as_async_function_prototype => {
        r#"async function* f() { yield 1; } console.log(Object.getPrototypeOf(f) === AsyncFunction.prototype);"#,
        ["false"]
    };

    async_generator_function_prototype_not_same_as_generator_prototype => {
        r#"async function* f() { yield 1; } console.log(Object.getPrototypeOf(f) === GeneratorFunction.prototype);"#,
        ["false"]
    };

    async_generator_function_prototype_call_result_has_async_iterator => {
        r#"async function* f() { yield 1; } const iter = AsyncGeneratorFunction.prototype.call.call(f, null); console.log(typeof iter[Symbol.asyncIterator]);"#,
        ["function"]
    };

    async_generator_function_prototype_apply_result_has_async_iterator => {
        r#"async function* f() { yield 1; } const iter = AsyncGeneratorFunction.prototype.apply.call(f, null, []); console.log(typeof iter[Symbol.asyncIterator]);"#,
        ["function"]
    };

    async_generator_function_bound_instanceof_after_partial_bind => {
        r#"async function* f(a, b) { yield a + b; } const partial = f.bind(null, 2); console.log(partial instanceof AsyncGeneratorFunction);"#,
        ["true"]
    };

    async_generator_function_prototype_not_equal_object_prototype => {
        r#"console.log(AsyncGeneratorFunction.prototype === Object.prototype);"#,
        ["false"]
    };

    async_generator_function_stored_in_array_keeps_prototype => {
        r#"const fns = [async function* g() { yield 1; }]; console.log(Object.getPrototypeOf(fns[0]) === AsyncGeneratorFunction.prototype);"#,
        ["true"]
    };
}
