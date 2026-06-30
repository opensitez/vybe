/// Function.prototype metadata — name, length, toString, Symbol.hasInstance, and per-kind naming rules.
crate::js_cases! {
    function_length_stops_at_first_default_parameter => {
        r#"function f(a, b = 1, c) {} console.log(f.length);"#,
        ["1"]
    };

    function_length_ignores_rest_parameter => {
        r#"function f(a, ...rest) {} console.log(f.length);"#,
        ["1"]
    };

    arrow_length_counts_params_without_default => {
        r#"const f = (a, b, c) => {}; console.log(f.length);"#,
        ["3"]
    };

    arrow_length_zero_when_first_param_has_default => {
        r#"const f = (a = 1, b) => {}; console.log(f.length);"#,
        ["0"]
    };

    arrow_length_zero_for_rest_only => {
        r#"const f = (...args) => args.length; console.log(f.length);"#,
        ["0"]
    };

    method_length_excludes_defaults_after_first => {
        r#"const obj = { m(a, b = 2, c) {} }; console.log(obj.m.length);"#,
        ["1"]
    };

    anonymous_assignment_infers_variable_name => {
        r#"const helper = function() {}; console.log(helper.name);"#,
        ["helper"]
    };

    object_property_method_infers_property_name => {
        r#"const obj = { run() {} }; console.log(obj.run.name);"#,
        ["run"]
    };

    object_property_function_value_infers_name => {
        r#"const obj = { run: function() {} }; console.log(obj.run.name);"#,
        ["run"]
    };

    arrow_assigned_to_variable_infers_name => {
        r#"const jump = () => {}; console.log(jump.name);"#,
        ["jump"]
    };

    arrow_assigned_to_property_infers_name => {
        r#"const obj = { hop: () => {} }; console.log(obj.hop.name);"#,
        ["hop"]
    };

    class_expression_infers_variable_name => {
        r#"const Widget = class {}; console.log(Widget.name);"#,
        ["Widget"]
    };

    class_declaration_uses_class_name => {
        r#"class Gadget {} console.log(Gadget.name);"#,
        ["Gadget"]
    };

    function_to_string_includes_function_keyword => {
        r#"function demo() {} console.log(Function.prototype.toString.call(demo).startsWith("function"));"#,
        ["true"]
    };

    arrow_to_string_includes_arrow_token => {
        r#"const a = () => 1; console.log(Function.prototype.toString.call(a).includes("=>"));"#,
        ["true"]
    };

    method_to_string_includes_method_name => {
        r#"const obj = { ping() { return 1; } }; console.log(Function.prototype.toString.call(obj.ping).includes("ping"));"#,
        ["true"]
    };

    bound_function_to_string_delegates_to_target => {
        r#"function orig() { return 0; } const b = orig.bind(null); console.log(Function.prototype.toString.call(b) === Function.prototype.toString.call(orig));"#,
        ["true"]
    };

    function_prototype_to_string_on_itself_is_native => {
        r#"const text = Function.prototype.toString.call(Function.prototype.toString); console.log(text.includes("toString"));"#,
        ["true"]
    };

    symbol_has_instance_on_function_constructor_exists => {
        r#"console.log(typeof Function[Symbol.hasInstance]);"#,
        ["function"]
    };

    instanceof_uses_function_has_instance_for_callable_objects => {
        r#"function Box() {} const b = new Box(); console.log(b instanceof Box);"#,
        ["true"]
    };

    custom_has_instance_can_force_true => {
        r#"
function Sentinel() {}
Object.defineProperty(Sentinel, Symbol.hasInstance, { value: () => true });
console.log({} instanceof Sentinel);
"#,
        ["true"]
    };

    custom_has_instance_can_force_false => {
        r#"
function Sentinel() {}
Object.defineProperty(Sentinel, Symbol.hasInstance, { value: () => false });
console.log(new Sentinel() instanceof Sentinel);
"#,
        ["false"]
    };

    function_name_property_is_configurable => {
        r#"function f() {} const desc = Object.getOwnPropertyDescriptor(f, "name"); console.log(desc.configurable);"#,
        ["true"]
    };

    function_length_property_is_configurable => {
        r#"function f(a) {} const desc = Object.getOwnPropertyDescriptor(f, "length"); console.log(desc.configurable);"#,
        ["true"]
    };

    function_name_is_non_enumerable => {
        r#"function f() {} console.log(f.propertyIsEnumerable("name"));"#,
        ["false"]
    };

    function_length_is_non_enumerable => {
        r#"function f(a) {} console.log(f.propertyIsEnumerable("length"));"#,
        ["false"]
    };

    reassigned_function_name_via_define_property => {
        r#"function f() {} Object.defineProperty(f, "name", { value: "renamed" }); console.log(f.name);"#,
        ["renamed"]
    };

    generator_function_name_from_variable => {
        r#"const gen = function* () { yield 1; }; console.log(gen.name);"#,
        ["gen"]
    };

    async_function_name_from_variable => {
        r#"const run = async function() { return 1; }; console.log(run.name);"#,
        ["run"]
    };

    async_arrow_name_from_variable => {
        r#"const run = async () => 1; console.log(run.name);"#,
        ["run"]
    };

    generator_declaration_has_name => {
        r#"function* produce() { yield 1; } console.log(produce.name);"#,
        ["produce"]
    };

    async_declaration_has_name => {
        r#"async function fetchValue() { return 1; } console.log(fetchValue.name);"#,
        ["fetchValue"]
    };

    function_with_unicode_name_in_expression => {
        r#"const π = function fn() {}; console.log(π.name);"#,
        ["fn"]
    };

    shorthand_method_named_symbol_key_has_empty_name => {
        r#"const s = Symbol("m"); const obj = { [s]() {} }; console.log(obj[s].name);"#,
        ["[s]"]
    };

    length_of_function_with_only_rest_is_zero => {
        r#"function only(...xs) {} console.log(only.length);"#,
        ["0"]
    };

    length_of_function_with_trailing_default_only => {
        r#"function tail(a, b, c = 1) {} console.log(tail.length);"#,
        ["2"]
    };

    to_string_on_native_function_reports_native => {
        r#"console.log(Function.prototype.toString.call(parseInt).includes("native") || Function.prototype.toString.call(parseInt).includes("function"));"#,
        ["true"]
    };

    function_prototype_name_is_empty => {
        r#"console.log(Function.prototype.name);"#,
        [""]
    };

    function_constructor_name_is_function => {
        r#"console.log(Function.name);"#,
        ["Function"]
    };

    empty_function_length_is_zero => {
        r#"function empty() {} console.log(empty.length);"#,
        ["0"]
    };

    single_param_arrow_length_is_one => {
        r#"const f = x => x; console.log(f.length);"#,
        ["1"]
    };

    method_in_class_has_class_context_name => {
        r#"class A { run() {} } console.log(A.prototype.run.name);"#,
        ["run"]
    };

    static_method_name_in_class => {
        r#"class A { static run() {} } console.log(A.run.name);"#,
        ["run"]
    };

    computed_method_name_uses_string_form => {
        r#"const obj = { ["move"]() {} }; console.log(obj.move.name);"#,
        ["move"]
    };

    function_expression_assigned_twice_keeps_initial_name => {
        r#"let a = function named() {}; a = function named() {}; console.log(a.name);"#,
        ["named"]
    };

    has_instance_on_arrow_is_function_builtin => {
        r#"const arrow = () => {}; console.log(Function[Symbol.hasInstance](arrow));"#,
        ["true"]
    };

    has_instance_rejects_plain_object_for_function => {
        r#"function F() {} console.log(({}) instanceof F);"#,
        ["false"]
    };

    has_instance_accepts_subclass_instance => {
        r#"class Base {} class Child extends Base {} console.log(new Child() instanceof Base);"#,
        ["true"]
    };

    function_to_string_contains_parameter_list => {
        r#"function pair(a, b) {} console.log(Function.prototype.toString.call(pair).includes("pair"));"#,
        ["true"]
    };

    async_method_to_string_contains_async_keyword => {
        r#"const obj = { async load() {} }; console.log(Function.prototype.toString.call(obj.load).includes("async"));"#,
        ["true"]
    };

    generator_method_to_string_contains_star => {
        r#"const obj = { *stream() { yield 1; } }; console.log(Function.prototype.toString.call(obj.stream).includes("*"));"#,
        ["true"]
    };

    bound_function_name_reflects_target_name => {
        r#"function target() {} const b = target.bind(null); console.log(b.name.includes("target"));"#,
        ["true"]
    };

    function_with_destructured_param_reduces_length => {
        r#"function f({ a }, b) {} console.log(f.length);"#,
        ["2"]
    };

    function_with_destructured_array_param_length => {
        r#"function f([a], b) {} console.log(f.length);"#,
        ["2"]
    };

    async_generator_function_name_from_declaration => {
        r#"async function* stream() { yield 1; } console.log(stream.name);"#,
        ["stream"]
    };

    function_prototype_length_is_zero => {
        r#"console.log(Function.prototype.length);"#,
        ["0"]
    };

    call_does_not_change_function_name => {
        r#"function labeled() {} labeled.call(null); console.log(labeled.name);"#,
        ["labeled"]
    };

    apply_does_not_change_function_length => {
        r#"function sized(a, b) {} sized.apply(null, [1, 2]); console.log(sized.length);"#,
        ["2"]
    };
}
