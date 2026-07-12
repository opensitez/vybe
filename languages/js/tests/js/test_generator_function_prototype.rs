// GeneratorFunction.prototype — generator metadata, prototype chain, and instanceof relations.
crate::js_cases! {
    generator_declaration_prototype_is_generator_function_prototype => {
        r#"function* run() { yield 1; } console.log(Object.getPrototypeOf(run) === GeneratorFunction.prototype);"#,
        ["true"]
    };

    generator_expression_prototype_is_generator_function_prototype => {
        r#"const run = function* () { yield 1; }; console.log(Object.getPrototypeOf(run) === GeneratorFunction.prototype);"#,
        ["true"]
    };

    named_generator_expression_prototype_is_generator_function_prototype => {
        r#"const outer = function* inner() { yield 1; }; console.log(Object.getPrototypeOf(outer) === GeneratorFunction.prototype);"#,
        ["true"]
    };

    generator_method_prototype_is_generator_function_prototype => {
        r#"const obj = { *m() { yield 1; } }; console.log(Object.getPrototypeOf(obj.m) === GeneratorFunction.prototype);"#,
        ["true"]
    };

    generator_class_method_prototype_is_generator_function_prototype => {
        r#"class S { *load() { yield 1; } } console.log(Object.getPrototypeOf(S.prototype.load) === GeneratorFunction.prototype);"#,
        ["true"]
    };

    generator_static_method_prototype_is_generator_function_prototype => {
        r#"class S { static *load() { yield 1; } } console.log(Object.getPrototypeOf(S.load) === GeneratorFunction.prototype);"#,
        ["true"]
    };

    generator_function_instanceof_generator_function => {
        r#"function* f() { yield 1; } console.log(f instanceof GeneratorFunction);"#,
        ["true"]
    };

    generator_expression_instanceof_generator_function => {
        r#"const f = function* () { yield 1; }; console.log(f instanceof GeneratorFunction);"#,
        ["true"]
    };

    generator_method_instanceof_generator_function => {
        r#"const obj = { *step() { yield 1; } }; console.log(obj.step instanceof GeneratorFunction);"#,
        ["true"]
    };

    generator_function_instanceof_function => {
        r#"function* f() { yield 1; } console.log(f instanceof Function);"#,
        ["true"]
    };

    generator_function_not_instanceof_async_function => {
        r#"function* f() { yield 1; } console.log(f instanceof AsyncFunction);"#,
        ["false"]
    };

    generator_function_not_instanceof_async_generator_function => {
        r#"function* f() { yield 1; } console.log(f instanceof AsyncGeneratorFunction);"#,
        ["false"]
    };

    regular_function_not_instanceof_generator_function => {
        r#"function f() {} console.log(f instanceof GeneratorFunction);"#,
        ["false"]
    };

    generator_function_prototype_prototype_is_function_prototype => {
        r#"console.log(Object.getPrototypeOf(GeneratorFunction.prototype) === Function.prototype);"#,
        ["true"]
    };

    generator_function_constructor_prototype_is_function_prototype => {
        r#"console.log(Object.getPrototypeOf(GeneratorFunction) === Function.prototype);"#,
        ["true"]
    };

    generator_function_constructor_name_is_generator_function => {
        r#"console.log(GeneratorFunction.name);"#,
        ["GeneratorFunction"]
    };

    generator_function_prototype_name_is_empty => {
        r#"console.log(GeneratorFunction.prototype.name);"#,
        [""]
    };

    generator_function_prototype_length_is_zero => {
        r#"console.log(GeneratorFunction.prototype.length);"#,
        ["0"]
    };

    generator_function_length_counts_simple_params => {
        r#"function* f(a, b, c) { yield 1; } console.log(f.length);"#,
        ["3"]
    };

    generator_function_length_stops_at_default => {
        r#"function* f(a, b = 1) { yield 1; } console.log(f.length);"#,
        ["1"]
    };

    generator_expression_length_with_rest_only_is_zero => {
        r#"const f = function* (...xs) { yield xs.length; }; console.log(f.length);"#,
        ["0"]
    };

    generator_method_length_excludes_defaults_after_first => {
        r#"const obj = { *m(a, b = 2, c) { yield 1; } }; console.log(obj.m.length);"#,
        ["1"]
    };

    generator_function_to_string_contains_function_star => {
        r#"function* ping() { yield 1; } console.log(Function.prototype.toString.call(ping).includes("function"));"#,
        ["true"]
    };

    generator_expression_to_string_contains_star_token => {
        r#"const f = function* () { yield 1; }; console.log(Function.prototype.toString.call(f).includes("*"));"#,
        ["true"]
    };

    generator_function_call_returns_iterator_object => {
        r#"function* value() { yield 4; } console.log(typeof value().next);"#,
        ["function"]
    };

    generator_function_apply_returns_iterator_object => {
        r#"function* value() { yield 4; } console.log(typeof value.apply(null, []).next);"#,
        ["function"]
    };

    generator_function_bind_returns_bound_generator => {
        r#"function* value() { yield this.n; } const b = value.bind({ n: 2 }); console.log(typeof b().next);"#,
        ["function"]
    };

    bound_generator_function_prototype_is_generator_function_prototype => {
        r#"function* value() { yield 1; } const b = value.bind(null); console.log(Object.getPrototypeOf(b) === GeneratorFunction.prototype);"#,
        ["true"]
    };

    bound_generator_function_instanceof_generator_function => {
        r#"function* value() { yield 1; } const b = value.bind(null); console.log(b instanceof GeneratorFunction);"#,
        ["true"]
    };

    // §27.3: generator functions DO have an own `prototype` (the object
    // their instances inherit from) — node-verified true. The old
    // expectation (false) confused generators with async functions.
    generator_function_has_no_own_prototype_property => {
        r#"function* f() { yield 1; } console.log(f.hasOwnProperty("prototype"));"#,
        ["true"]
    };

    // §27.3: a generator expression's `prototype` is an ordinary object,
    // not undefined (node-verified false).
    generator_expression_prototype_property_is_undefined => {
        r#"const f = function* () { yield 1; }; console.log(f.prototype === undefined);"#,
        ["false"]
    };

    generator_function_constructor_is_generator_function => {
        r#"console.log(function* f() { yield 1; }.constructor === GeneratorFunction);"#,
        ["true"]
    };

    generator_function_prototype_has_call_method => {
        r#"console.log(typeof GeneratorFunction.prototype.call);"#,
        ["function"]
    };

    generator_function_prototype_has_apply_method => {
        r#"console.log(typeof GeneratorFunction.prototype.apply);"#,
        ["function"]
    };

    generator_function_prototype_has_bind_method => {
        r#"console.log(typeof GeneratorFunction.prototype.bind);"#,
        ["function"]
    };

    generator_function_prototype_has_to_string => {
        r#"console.log(typeof GeneratorFunction.prototype.toString);"#,
        ["function"]
    };

    generator_function_prototype_inherits_from_function_prototype => {
        r#"console.log("call" in GeneratorFunction.prototype && "apply" in GeneratorFunction.prototype);"#,
        ["true"]
    };

    generator_function_is_not_same_as_generator_function_prototype => {
        r#"function* f() { yield 1; } console.log(f === GeneratorFunction.prototype);"#,
        ["false"]
    };

    generator_function_prototype_is_not_generator_function_instance => {
        r#"function* f() { yield 1; } console.log(GeneratorFunction.prototype instanceof GeneratorFunction);"#,
        ["false"]
    };

    generator_iife_call_returns_iterator_not_function => {
        r#"const iter = (function* () { yield 1; })(); console.log(typeof iter.next);"#,
        ["function"]
    };

    generator_method_name_in_object_literal => {
        r#"const api = { *fetch() { yield 1; } }; console.log(api.fetch.name);"#,
        ["fetch"]
    };

    named_generator_expression_preserves_inner_name => {
        r#"const outer = function* inner() { yield 1; }; console.log(outer.name);"#,
        ["inner"]
    };

    generator_expression_name_from_property_assignment => {
        r#"const tools = { maker: function* () { yield 1; } }; console.log(tools.maker.name);"#,
        ["maker"]
    };

    generator_method_extracted_preserves_prototype => {
        r#"const src = { *emit() { yield 1; } }; const bare = src.emit; console.log(Object.getPrototypeOf(bare) === GeneratorFunction.prototype);"#,
        ["true"]
    };

    generator_function_call_with_this_on_regular_generator => {
        r#"function* read() { yield this.v; } const iter = read.call({ v: 9 }); console.log(iter.next().value);"#,
        ["9"]
    };

    generator_function_apply_with_args_yields_values => {
        r#"function* add(a, b) { yield a + b; } console.log(add.apply(null, [2, 3]).next().value);"#,
        ["5"]
    };

    generator_function_bind_partial_preserves_generator_prototype => {
        r#"function* add(a, b) { yield a + b; } const plusOne = add.bind(null, 1); console.log(Object.getPrototypeOf(plusOne) === GeneratorFunction.prototype);"#,
        ["true"]
    };

    generator_function_bind_length_reduced => {
        r#"function* f(a, b) { yield a + b; } console.log(f.bind(null, 1).length);"#,
        ["1"]
    };

    generator_function_instanceof_object_is_true => {
        r#"function* f() { yield 1; } console.log(f instanceof Object);"#,
        ["true"]
    };

    generator_function_prototype_not_equal_function_prototype => {
        r#"console.log(GeneratorFunction.prototype === Function.prototype);"#,
        ["false"]
    };

    // §27.3.3: %GeneratorFunction.prototype% is an ordinary object, not a
    // function (node-verified "object").
    generator_function_prototype_typeof_is_function => {
        r#"console.log(typeof GeneratorFunction.prototype);"#,
        ["object"]
    };

    generator_function_prototype_constructor_is_generator_function => {
        r#"console.log(GeneratorFunction.prototype.constructor === GeneratorFunction);"#,
        ["true"]
    };

    generator_function_from_new_function_is_not_generator_function => {
        r#"const f = new Function("return 1;"); console.log(f instanceof GeneratorFunction);"#,
        ["false"]
    };

    generator_function_prototype_to_string_returns_string => {
        r#"console.log(typeof GeneratorFunction.prototype.toString.call(function* () { yield 1; }));"#,
        ["string"]
    };

    generator_function_has_configurable_name => {
        r#"function* f() { yield 1; } console.log(Object.getOwnPropertyDescriptor(f, "name").configurable);"#,
        ["true"]
    };

    generator_function_has_configurable_length => {
        r#"function* f(a) { yield 1; } console.log(Object.getOwnPropertyDescriptor(f, "length").configurable);"#,
        ["true"]
    };

    generator_function_symbol_has_instance_via_function => {
        r#"function* F() { yield 1; } console.log(F[Symbol.hasInstance] === Function[Symbol.hasInstance]);"#,
        ["true"]
    };

    generator_assigned_to_property_infers_name => {
        r#"const obj = { task: function* () { yield 1; } }; console.log(obj.task.name);"#,
        ["task"]
    };

    generator_function_prototype_call_invokes_and_returns_iterator => {
        r#"function* f() { yield 1; } console.log(typeof GeneratorFunction.prototype.call.call(f, null).next);"#,
        ["function"]
    };

    generator_function_prototype_apply_invokes_and_returns_iterator => {
        r#"function* f() { yield 1; } console.log(typeof GeneratorFunction.prototype.apply.call(f, null, []).next);"#,
        ["function"]
    };

    generator_function_prototype_bind_creates_generator_callable => {
        r#"function* f() { yield 1; } const b = GeneratorFunction.prototype.bind.call(f, null); console.log(b instanceof GeneratorFunction);"#,
        ["true"]
    };

    object_get_prototype_of_generator_is_not_function_prototype => {
        r#"function* f() { yield 1; } console.log(Object.getPrototypeOf(f) === Function.prototype);"#,
        ["false"]
    };

    generator_function_extends_function_prototype_chain => {
        r#"function* f() { yield 1; } console.log(Function.prototype.isPrototypeOf(f));"#,
        ["true"]
    };

    generator_function_prototype_is_prototype_of_generator_instances => {
        r#"function* f() { yield 1; } console.log(GeneratorFunction.prototype.isPrototypeOf(f));"#,
        ["true"]
    };

    function_prototype_is_not_direct_prototype_of_generator => {
        r#"function* f() { yield 1; } console.log(Object.getPrototypeOf(f) === Function.prototype);"#,
        ["false"]
    };

    generator_class_instance_method_instanceof_generator_function => {
        r#"class Box { *values() { yield 1; } } const b = new Box(); console.log(b.values instanceof GeneratorFunction);"#,
        ["true"]
    };

    generator_static_class_method_instanceof_generator_function => {
        r#"class Box { static *values() { yield 1; } } console.log(Box.values instanceof GeneratorFunction);"#,
        ["true"]
    };

    generator_computed_method_name_preserves_name => {
        r#"const key = "stream"; const obj = { *[key]() { yield 1; } }; console.log(obj[key].name);"#,
        ["stream"]
    };

    generator_symbol_method_name_is_symbol_description => {
        r#"const sym = Symbol("iter"); const obj = { *[sym]() { yield 1; } }; console.log(obj[sym].name);"#,
        ["[iter]"]
    };

    generator_function_length_with_only_rest_is_zero => {
        r#"function* only(...xs) { yield xs.length; } console.log(only.length);"#,
        ["0"]
    };

    generator_function_length_with_trailing_default_only => {
        r#"function* tail(a, b, c = 1) { yield 1; } console.log(tail.length);"#,
        ["2"]
    };

    empty_generator_function_length_is_zero => {
        r#"function* empty() { yield 1; } console.log(empty.length);"#,
        ["0"]
    };

    generator_function_length_is_non_enumerable => {
        r#"function* f(a) { yield 1; } console.log(f.propertyIsEnumerable("length"));"#,
        ["false"]
    };

    generator_function_name_is_non_enumerable => {
        r#"function* f() { yield 1; } console.log(f.propertyIsEnumerable("name"));"#,
        ["false"]
    };

    generator_function_call_result_is_not_generator_function => {
        r#"function* f() { yield 1; } console.log(f() instanceof GeneratorFunction);"#,
        ["false"]
    };

    generator_function_call_result_has_symbol_iterator => {
        r#"function* f() { yield 1; } console.log(typeof f()[Symbol.iterator]);"#,
        ["function"]
    };

    generator_bound_call_preserves_yield_behavior => {
        r#"function* f(x) { yield x * 2; } const doubled = f.bind(null, 3); console.log(doubled().next().value);"#,
        ["6"]
    };

    generator_function_prototype_is_not_enumerable_on_global => {
        r#"console.log(Object.prototype.propertyIsEnumerable.call(globalThis, "GeneratorFunction"));"#,
        ["false"]
    };

    generator_function_constructor_is_function => {
        r#"console.log(GeneratorFunction instanceof Function);"#,
        ["true"]
    };

    generator_function_constructor_length_is_one => {
        r#"console.log(GeneratorFunction.length);"#,
        ["1"]
    };

    generator_function_prototype_call_passes_arguments => {
        r#"function* pick(a, b) { yield b; } console.log(GeneratorFunction.prototype.call.call(pick, null, 1, 9).next().value);"#,
        ["9"]
    };

    generator_function_prototype_apply_spreads_array_args => {
        r#"function* pick(a, b) { yield a + b; } console.log(GeneratorFunction.prototype.apply.call(pick, null, [4, 5]).next().value);"#,
        ["9"]
    };

    // §10.2.10 SetFunctionLength: a destructured pattern counts as ONE
    // formal parameter — node-verified 2. The old expectation (1) treated
    // the pattern as if it reduced length.
    generator_function_with_destructured_param_reduces_length => {
        r#"function* f({ a }, b) { yield a + b; } console.log(f.length);"#,
        ["2"]
    };

    // §10.2.10 SetFunctionLength: a destructured pattern counts as ONE
    // formal parameter — node-verified 2. The old expectation (1) treated
    // the pattern as if it reduced length.
    generator_function_with_destructured_array_param_length => {
        r#"function* f([a], b) { yield a + b; } console.log(f.length);"#,
        ["2"]
    };

    generator_function_strict_mode_preserves_prototype => {
        r#""use strict"; function* f() { yield 1; } console.log(Object.getPrototypeOf(f) === GeneratorFunction.prototype);"#,
        ["true"]
    };

    generator_expression_strict_mode_instanceof => {
        r#""use strict"; const f = function* () { yield 1; }; console.log(f instanceof GeneratorFunction);"#,
        ["true"]
    };

    generator_function_is_not_constructable_via_new => {
        r#"function* f() { yield 1; } try { new f(); console.log("constructed"); } catch (e) { console.log(e instanceof TypeError); }"#,
        ["true"]
    };

    generator_function_prototype_value_of_inherits_from_function => {
        r#"console.log(typeof GeneratorFunction.prototype.valueOf);"#,
        ["function"]
    };

    generator_function_to_string_on_class_method_includes_star => {
        r#"class G { *run() { yield 1; } } console.log(Function.prototype.toString.call(G.prototype.run).includes("*"));"#,
        ["true"]
    };

    generator_function_call_with_null_this_in_strict => {
        r#""use strict"; function* f() { yield this; } console.log(f.call(null).next().value === null);"#,
        ["true"]
    };

    generator_function_call_with_object_this_in_sloppy => {
        r#"function* f() { yield this.tag; } console.log(f.call({ tag: "hit" }).next().value);"#,
        ["hit"]
    };

    generator_function_bind_with_this_forwards_receiver => {
        r#"function* f() { yield this.n; } const b = f.bind({ n: 7 }); console.log(b().next().value);"#,
        ["7"]
    };

    generator_function_prototype_chain_reaches_object_prototype => {
        r#"console.log(Object.prototype.isPrototypeOf(GeneratorFunction.prototype));"#,
        ["true"]
    };

    generator_function_instance_prototype_chain_reaches_object => {
        r#"function* f() { yield 1; } console.log(Object.prototype.isPrototypeOf(f));"#,
        ["true"]
    };

    async_generator_not_instanceof_generator_function => {
        r#"async function* g() { yield 1; } console.log(g instanceof GeneratorFunction);"#,
        ["false"]
    };

    arrow_function_not_instanceof_generator_function => {
        r#"const f = () => 1; console.log(f instanceof GeneratorFunction);"#,
        ["false"]
    };

    // §10.4.1.3: the bound generator carries (1, 2); the third parameter
    // comes from the call — `partial()` alone yields NaN in node since
    // `c` is undefined. Pass the remaining argument.
    generator_function_prototype_bind_partial_application => {
        r#"function* sum(a, b, c) { yield a + b + c; } const partial = GeneratorFunction.prototype.bind.call(sum, null, 1, 2); console.log(partial(3).next().value);"#,
        ["6"]
    };

    generator_function_multiple_yields_via_call => {
        r#"function* seq() { yield 1; yield 2; } const it = seq(); console.log(it.next().value + it.next().value);"#,
        ["3"]
    };

    generator_function_return_value_available_after_exhaustion => {
        r#"function* fin() { yield 1; return 9; } const it = fin(); it.next(); console.log(it.next().value);"#,
        ["9"]
    };

    generator_function_prototype_is_function_object => {
        r#"console.log(GeneratorFunction.prototype instanceof Object);"#,
        ["true"]
    };

    generator_function_global_binding_is_function => {
        r#"console.log(typeof GeneratorFunction);"#,
        ["function"]
    };

    generator_function_prototype_differs_from_async_generator_prototype => {
        r#"console.log(GeneratorFunction.prototype === AsyncGeneratorFunction.prototype);"#,
        ["false"]
    };

    generator_function_constructor_differs_from_async_generator_constructor => {
        r#"console.log(GeneratorFunction === AsyncGeneratorFunction);"#,
        ["false"]
    };

    generator_function_prototype_differs_from_async_function_prototype => {
        r#"console.log(GeneratorFunction.prototype === AsyncFunction.prototype);"#,
        ["false"]
    };

    generator_nested_declaration_keeps_outer_name => {
        r#"function outer() { function* inner() { yield 1; } return inner.name; } console.log(outer());"#,
        ["inner"]
    };

    generator_function_call_result_next_done_false_initially => {
        r#"function* f() { yield 1; } console.log(f().next().done);"#,
        ["false"]
    };

    generator_function_call_result_next_done_true_after_return => {
        r#"function* f() { return 1; } console.log(f().next().done);"#,
        ["true"]
    };

    generator_function_prototype_not_equal_object_prototype => {
        r#"console.log(GeneratorFunction.prototype === Object.prototype);"#,
        ["false"]
    };

    generator_function_stored_in_array_keeps_prototype => {
        r#"const fns = [function* g() { yield 1; }]; console.log(Object.getPrototypeOf(fns[0]) === GeneratorFunction.prototype);"#,
        ["true"]
    };
}
