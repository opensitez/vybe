// Function.prototype — call, apply, bind semantics and arrow vs function vs method prototype differences.
crate::js_cases! {
    call_with_object_receiver_reads_property => {
        r#"function greet() { return this.name; } console.log(greet.call({ name: "Ada" }));"#,
        ["Ada"]
    };

    call_with_null_this_sloppy_returns_global_object => {
        r#"function f() { return typeof this; } console.log(f.call(null) === "object" || f.call(null) === "undefined");"#,
        ["true"]
    };

    call_with_null_this_strict_returns_null => {
        r#""use strict"; function f() { return this; } console.log(f.call(null) === null);"#,
        ["true"]
    };

    call_with_undefined_this_strict_returns_undefined => {
        r#""use strict"; function f() { return this; } console.log(f.call(undefined) === undefined);"#,
        ["true"]
    };

    call_passes_single_argument_to_target => {
        r#"function add(a, b) { return a + b; } console.log(add.call(null, 5));"#,
        ["NaN"]
    };

    call_passes_three_arguments_in_order => {
        r#"function sum(a, b, c) { return a + b + c; } console.log(sum.call(null, 1, 2, 3));"#,
        ["6"]
    };

    call_with_only_this_argument_uses_undefined_params => {
        r#"function pair(a, b) { return [a, b]; } console.log(JSON.stringify(pair.call(null)));"#,
        ["[null,null]"]
    };

    call_borrowed_array_join_on_array_like_object => {
        r#"const like = { 0: "a", 1: "b", length: 2 }; console.log(Array.prototype.join.call(like, "-"));"#,
        ["a-b"]
    };

    call_borrowed_array_map_on_array_like_object => {
        r#"const like = { 0: 1, 1: 2, length: 2 }; console.log(JSON.stringify(Array.prototype.map.call(like, x => x * 10)));"#,
        ["[10,20]"]
    };

    call_on_arrow_function_ignores_this_argument => {
        r#"const outer = { x: 1 }; const arrow = () => this; console.log(arrow.call({ x: 9 }) === arrow.call(outer));"#,
        ["true"]
    };

    call_extracted_method_without_receiver_loses_this => {
        r#"const obj = { n: 3, read() { return this.n; } }; const bare = obj.read; try { bare.call(null); console.log("ok"); } catch (e) { console.log(e instanceof TypeError); }"#,
        ["true"]
    };

    call_with_explicit_this_mutates_receiver => {
        r#"const bag = { count: 0 }; function inc() { this.count++; } inc.call(bag); inc.call(bag); console.log(bag.count);"#,
        ["2"]
    };

    call_borrowed_string_slice => {
        r#"console.log(String.prototype.slice.call("hello", 1, 4));"#,
        ["ell"]
    };

    call_borrowed_object_hasownproperty => {
        r#"const obj = { a: 1 }; console.log(Object.prototype.hasOwnProperty.call(obj, "a"));"#,
        ["true"]
    };

    call_on_function_prototype_call_invokes_target => {
        r#"function target(v) { return v * 2; } console.log(Function.prototype.call.call(target, null, 4));"#,
        ["8"]
    };

    call_with_number_this_coerces_in_sloppy_mode => {
        r#"function tag() { return Object.prototype.toString.call(this); } console.log(tag.call(42).includes("Number"));"#,
        ["true"]
    };

    call_with_string_this_coerces_in_sloppy_mode => {
        r#"function tag() { return Object.prototype.toString.call(this); } console.log(tag.call("hi").includes("String"));"#,
        ["true"]
    };

    call_extra_arguments_beyond_formal_params_are_ignored => {
        r#"function takeOne(a) { return a; } console.log(takeOne.call(null, 1, 2, 3));"#,
        ["1"]
    };

    call_with_symbol_property_on_this => {
        r#"const key = Symbol("k"); const obj = { [key]: 7 }; function read() { return this[key]; } console.log(read.call(obj));"#,
        ["7"]
    };

    call_strict_function_with_object_this_keeps_identity => {
        r#""use strict"; function id() { return this; } const o = {}; console.log(id.call(o) === o);"#,
        ["true"]
    };

    apply_with_empty_array_passes_no_args => {
        r#"function len() { return arguments.length; } console.log(apply_len()); function apply_len() { return len.apply(null, []); }"#,
        ["0"]
    };

    apply_with_three_element_array_spreads_args => {
        r#"function sum(a, b, c) { return a + b + c; } console.log(sum.apply(null, [4, 5, 6]));"#,
        ["15"]
    };

    apply_with_array_like_object_with_length => {
        r#"function first(a) { return a; } const like = { 0: "z", length: 1 }; console.log(first.apply(null, like));"#,
        ["z"]
    };

    apply_with_null_this_in_strict_returns_null => {
        r#""use strict"; function f() { return this; } console.log(f.apply(null, []) === null);"#,
        ["true"]
    };

    apply_on_arrow_ignores_this_argument => {
        r#"const a = () => 1; const b = () => 1; console.log(a.apply({ x: 1 }, []) === b.apply({ x: 2 }, []));"#,
        ["true"]
    };

    apply_math_max_with_number_array => {
        r#"console.log(Math.max.apply(null, [3, 9, 2]));"#,
        ["9"]
    };

    apply_array_push_borrowed_on_plain_object => {
        r#"const obj = { length: 0 }; Array.prototype.push.apply(obj, ["a", "b"]); console.log(obj.length);"#,
        ["2"]
    };

    apply_with_single_element_array => {
        r#"function double(x) { return x * 2; } console.log(double.apply(null, [6]));"#,
        ["12"]
    };

    apply_strict_undefined_this => {
        r#""use strict"; function f() { return this; } console.log(f.apply(undefined, []) === undefined);"#,
        ["true"]
    };

    apply_passes_array_elements_not_array_itself => {
        r#"function isArray(a) { return Array.isArray(a); } console.log(isArray.apply(null, [[1]]));"#,
        ["true"]
    };

    apply_on_method_with_context_object => {
        r#"const ctx = { factor: 3 }; function scale(n) { return n * this.factor; } console.log(scale.apply(ctx, [4]));"#,
        ["12"]
    };

    apply_with_sparse_array_preserves_holes_as_undefined => {
        r#"function pair(a, b) { return a === undefined && b === undefined; } const sparse = [1, , 3]; console.log(pair.apply(null, sparse));"#,
        ["false"]
    };

    apply_variadic_from_runtime_array => {
        r#"function tag() { return Array.from(arguments).join(","); } console.log(tag.apply(null, ["x", "y"]));"#,
        ["x,y"]
    };

    apply_borrowed_object_keys_on_plain_object => {
        r#"const obj = { a: 1, b: 2 }; console.log(Object.keys.apply(obj, []).sort().join(","));"#,
        ["a,b"]
    };

    apply_on_bound_target_still_invokes => {
        r#"function add(a, b) { return a + b; } const plusOne = add.bind(null, 1); console.log(plusOne.apply(null, [2]));"#,
        ["3"]
    };

    bind_creates_distinct_function_object => {
        r#"function f() {} const b = f.bind(null); console.log(b === f);"#,
        ["false"]
    };

    bind_partial_first_argument => {
        r#"function sub(a, b) { return a - b; } const fromTen = sub.bind(null, 10); console.log(fromTen(3));"#,
        ["7"]
    };

    bind_partial_two_arguments => {
        r#"function mul(a, b, c) { return a * b * c; } const timesSix = mul.bind(null, 2, 3); console.log(timesSix(4));"#,
        ["24"]
    };

    bind_without_partial_args_only_fixes_this => {
        r#"function read() { return this.v; } const bound = read.bind({ v: "ok" }); console.log(bound());"#,
        ["ok"]
    };

    bind_length_decreases_by_bound_arg_count => {
        r#"function f(a, b, c) {} console.log(f.bind(null, 1).length);"#,
        ["2"]
    };

    bind_length_zero_when_all_params_bound => {
        r#"function f(a, b) {} console.log(f.bind(null, 1, 2).length);"#,
        ["0"]
    };

    bind_name_includes_bound_prefix => {
        r#"function named() {} const b = named.bind(null); console.log(b.name.includes("bound"));"#,
        ["true"]
    };

    bind_second_bind_composes_partial_args => {
        r#"function concat(a, b, c) { return "" + a + b + c; } const step = concat.bind(null, "a"); const done = step.bind(null, "b"); console.log(done("c"));"#,
        ["abc"]
    };

    bind_cannot_override_fixed_this_via_call => {
        r#"function id() { return this.tag; } const fixed = id.bind({ tag: "one" }); console.log(fixed.call({ tag: "two" }));"#,
        ["one"]
    };

    bind_null_this_with_partial_args => {
        r#"function add(a, b) { return a + b; } console.log(add.bind(null, 5)(7));"#,
        ["12"]
    };

    bind_extracted_instance_method_preserves_this => {
        r#"const counter = { n: 0, tick() { this.n++; return this.n; } }; const tick = counter.tick.bind(counter); console.log(tick());"#,
        ["1"]
    };

    bind_zero_arg_function_has_length_zero => {
        r#"function noop() { return 1; } console.log(noop.bind(null).length);"#,
        ["0"]
    };

    bind_more_partial_args_than_formals_still_callable => {
        r#"function take(a) { return a; } const fixed = take.bind(null, 1, 2, 3); console.log(fixed());"#,
        ["1"]
    };

    bind_target_prototype_is_preserved_on_bound_function => {
        r#"function decl() {} const b = decl.bind(null); console.log(Object.getPrototypeOf(b) === Function.prototype);"#,
        ["true"]
    };

    bind_arrow_function_prototype_is_function_prototype => {
        r#"const arrow = () => 1; const b = arrow.bind(null); console.log(Object.getPrototypeOf(b) === Function.prototype);"#,
        ["true"]
    };

    declaration_get_prototype_of_is_function_prototype => {
        r#"function decl() {} console.log(Object.getPrototypeOf(decl) === Function.prototype);"#,
        ["true"]
    };

    expression_get_prototype_of_is_function_prototype => {
        r#"const expr = function() {}; console.log(Object.getPrototypeOf(expr) === Function.prototype);"#,
        ["true"]
    };

    arrow_get_prototype_of_is_function_prototype => {
        r#"const arrow = () => {}; console.log(Object.getPrototypeOf(arrow) === Function.prototype);"#,
        ["true"]
    };

    concise_method_get_prototype_of_is_function_prototype => {
        r#"const obj = { m() {} }; console.log(Object.getPrototypeOf(obj.m) === Function.prototype);"#,
        ["true"]
    };

    declaration_instanceof_function => {
        r#"function decl() {} console.log(decl instanceof Function);"#,
        ["true"]
    };

    arrow_instanceof_function => {
        r#"const arrow = () => {}; console.log(arrow instanceof Function);"#,
        ["true"]
    };

    method_instanceof_function => {
        r#"const obj = { m() {} }; console.log(obj.m instanceof Function);"#,
        ["true"]
    };

    arrow_has_no_prototype_property => {
        r#"const arrow = () => {}; console.log(arrow.prototype === undefined);"#,
        ["true"]
    };

    function_declaration_has_prototype_object => {
        r#"function ctor() {} console.log(typeof ctor.prototype === "object");"#,
        ["true"]
    };

    arrow_prototype_is_undefined_not_null => {
        r#"const arrow = () => {}; console.log(arrow.hasOwnProperty("prototype"));"#,
        ["false"]
    };

    concise_method_has_no_own_prototype_property => {
        r#"const obj = { m() {} }; console.log(obj.m.hasOwnProperty("prototype"));"#,
        ["false"]
    };

    function_prototype_exposes_call_as_function => {
        r#"console.log(typeof Function.prototype.call);"#,
        ["function"]
    };

    function_prototype_exposes_apply_as_function => {
        r#"console.log(typeof Function.prototype.apply);"#,
        ["function"]
    };

    function_prototype_exposes_bind_as_function => {
        r#"console.log(typeof Function.prototype.bind);"#,
        ["function"]
    };

    function_prototype_exposes_to_string => {
        r#"console.log(typeof Function.prototype.toString);"#,
        ["function"]
    };

    call_bound_function_with_extra_args_appended => {
        r#"function pair(a, b) { return a + b; } const one = pair.bind(null, 1); console.log(one(2, 99));"#,
        ["3"]
    };

    apply_on_function_with_default_param_uses_passed_value => {
        r#"function f(a, b = 10) { return a + b; } console.log(f.apply(null, [2, 3]));"#,
        ["5"]
    };

    bind_on_class_method_preserves_home_object_this_when_bound => {
        r#"class Box { constructor() { this.v = 1; } read() { return this.v; } } const b = new Box(); const read = b.read.bind(b); console.log(read());"#,
        ["1"]
    };
}
