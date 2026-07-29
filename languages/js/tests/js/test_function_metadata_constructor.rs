crate::js_cases! {
    function_constructor_builds_callable_function => {
        r#"
const add = new Function("a", "b", "return a + b;");
console.log(add(2, 3));
"#,
        ["5"]
    };

    function_constructor_without_new_also_creates_function => {
        r#"
const double = Function("x", "return x * 2;");
console.log(double(9));
"#,
        ["18"]
    };

    function_constructor_can_return_object_literal => {
        r#"
const make = new Function("return { x: 1, y: 2 };");
const value = make();
console.log(value.x + value.y);
"#,
        ["3"]
    };

    function_constructor_does_not_capture_local_scope => {
        r#"
let hidden = 42;
const fn = new Function("return typeof hidden;");
console.log(fn());
"#,
        ["undefined"]
    };

    function_constructor_can_read_global_this => {
        r#"
globalThis.dynamicValue = 7;
const fn = new Function("return globalThis.dynamicValue;");
console.log(fn());
delete globalThis.dynamicValue;
"#,
        ["7"]
    };

    function_constructor_name_is_anonymous => {
        r#"
const fn = new Function("");
console.log(fn.name);
"#,
        ["anonymous"]
    };

    function_constructor_length_counts_declared_parameters => {
        r#"
const fn = new Function("a", "b", "c", "return 0;");
console.log(fn.length);
"#,
        ["3"]
    };

    function_constructor_result_is_instance_of_function => {
        r#"
const fn = new Function("return 1;");
console.log(fn instanceof Function);
"#,
        ["true"]
    };

    function_declaration_length_stops_before_default_parameter => {
        r#"
function sample(a, b = 1, c) {}
console.log(sample.length);
"#,
        ["1"]
    };

    function_declaration_length_ignores_rest_parameter => {
        r#"
function sample(a, ...rest) {}
console.log(sample.length);
"#,
        ["1"]
    };

    arrow_function_length_counts_simple_parameters => {
        r#"
const fn = (a, b) => a + b;
console.log(fn.length);
"#,
        ["2"]
    };

    arrow_function_length_is_zero_when_first_parameter_has_default => {
        r#"
const fn = (a = 1, b) => a + b;
console.log(fn.length);
"#,
        ["0"]
    };

    arrow_function_length_is_zero_for_rest_parameter_only => {
        r#"
const fn = (...args) => args.length;
console.log(fn.length);
"#,
        ["0"]
    };

    anonymous_function_assigned_to_variable_infers_variable_name => {
        r#"
const named = function() {};
console.log(named.name);
"#,
        ["named"]
    };

    named_function_expression_preserves_inner_name => {
        r#"
const outer = function inner() {};
console.log(outer.name);
"#,
        ["inner"]
    };

    anonymous_function_assigned_to_object_property_infers_property_name => {
        r#"
const obj = { method: function() {} };
console.log(obj.method.name);
"#,
        ["method"]
    };

    concise_method_definition_uses_property_name_for_name => {
        r#"
const obj = { speak() {} };
console.log(obj.speak.name);
"#,
        ["speak"]
    };

    getter_name_includes_get_prefix => {
        r#"
const obj = {
  get size() { return 1; }
};
console.log(Object.getOwnPropertyDescriptor(obj, "size").get.name);
"#,
        ["get size"]
    };

    setter_name_includes_set_prefix => {
        r#"
const obj = {
  set size(value) {}
};
console.log(Object.getOwnPropertyDescriptor(obj, "size").set.name);
"#,
        ["set size"]
    };

    class_expression_assigned_to_variable_infers_name => {
        r#"
const Box = class {};
console.log(Box.name);
"#,
        ["Box"]
    };

    arguments_callee_references_current_non_strict_function => {
        r#"
const fn = function() {
  return fn === arguments.callee;
};
console.log(fn());
"#,
        ["true"]
    };

    function_bind_name_prefixes_bound => {
        r#"
function target() {}
console.log(target.bind(null).name);
"#,
        ["bound target"]
    };
}

