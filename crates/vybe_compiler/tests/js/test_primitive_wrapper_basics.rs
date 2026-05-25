crate::js_cases! {
    new_number_has_object_type => {
        r#"
console.log(typeof new Number(5));
"#,
        ["object"]
    };

    new_string_has_object_type => {
        r#"
console.log(typeof new String("hi"));
"#,
        ["object"]
    };

    new_boolean_has_object_type => {
        r#"
console.log(typeof new Boolean(false));
"#,
        ["object"]
    };

    number_wrapper_valueof_returns_primitive_number => {
        r#"
console.log(new Number(5).valueOf());
"#,
        ["5"]
    };

    string_wrapper_valueof_returns_primitive_string => {
        r#"
console.log(new String("hi").valueOf());
"#,
        ["hi"]
    };

    boolean_wrapper_valueof_returns_primitive_boolean => {
        r#"
console.log(new Boolean(false).valueOf());
"#,
        ["false"]
    };

    number_wrapper_to_string_uses_number_representation => {
        r#"
console.log(new Number(255).toString(16));
"#,
        ["ff"]
    };

    string_wrapper_to_string_uses_wrapped_text => {
        r#"
console.log(new String("hello").toString());
"#,
        ["hello"]
    };

    boolean_wrapper_to_string_uses_wrapped_boolean => {
        r#"
console.log(new Boolean(true).toString());
"#,
        ["true"]
    };

    number_wrapper_loose_equality_matches_same_primitive => {
        r#"
console.log(new Number(5) == 5);
"#,
        ["true"]
    };

    string_wrapper_loose_equality_matches_same_primitive => {
        r#"
console.log(new String("hi") == "hi");
"#,
        ["true"]
    };

    boolean_wrapper_loose_equality_matches_same_primitive => {
        r#"
console.log(new Boolean(false) == false);
"#,
        ["true"]
    };

    number_wrapper_strict_equality_differs_from_primitive => {
        r#"
console.log(new Number(5) === 5);
"#,
        ["false"]
    };

    string_wrapper_strict_equality_differs_from_primitive => {
        r#"
console.log(new String("hi") === "hi");
"#,
        ["false"]
    };

    boolean_wrapper_strict_equality_differs_from_primitive => {
        r#"
console.log(new Boolean(false) === false);
"#,
        ["false"]
    };

    boolean_wrapper_object_is_truthy_even_when_wrapping_false => {
        r#"
if (new Boolean(false)) {
  console.log("truthy");
} else {
  console.log("falsey");
}
"#,
        ["truthy"]
    };

    number_wrapper_addition_uses_primitive_value => {
        r#"
console.log(new Number(5) + 1);
"#,
        ["6"]
    };

    string_wrapper_concatenation_uses_primitive_value => {
        r#"
console.log(new String("a") + "b");
"#,
        ["ab"]
    };

    wrapper_objects_are_instances_of_their_constructors => {
        r#"
console.log(new Number(1) instanceof Number);
console.log(new String("x") instanceof String);
console.log(new Boolean(true) instanceof Boolean);
"#,
        ["true", "true", "true"]
    };

    primitive_constructor_property_points_to_builtin_constructor => {
        r#"
console.log("hi".constructor === String);
console.log((42).constructor === Number);
console.log(true.constructor === Boolean);
"#,
        ["true", "true", "true"]
    };

    primitive_valueof_methods_return_same_primitive => {
        r#"
console.log("hi".valueOf());
console.log((3.14).valueOf());
console.log(true.valueOf());
"#,
        ["hi", "3.14", "true"]
    };

    object_prototype_to_string_reports_wrapper_tags => {
        r#"
console.log(Object.prototype.toString.call(new Number(1)));
console.log(Object.prototype.toString.call(new String("x")));
console.log(Object.prototype.toString.call(new Boolean(false)));
"#,
        ["[object Number]", "[object String]", "[object Boolean]"]
    };

    string_wrapper_exposes_index_keys_for_characters => {
        r#"
console.log(Object.keys(new String("hi")).join(","));
"#,
        ["0,1"]
    };

    string_wrapper_length_matches_wrapped_text_length => {
        r#"
console.log(new String("hello").length);
"#,
        ["5"]
    };

    wrapper_objects_can_store_expando_properties => {
        r#"
const s = new String("hi");
s.extra = 1;
const n = new Number(2);
n.tag = "x";
console.log(s.extra);
console.log(n.tag);
"#,
        ["1", "x"]
    };

    distinct_string_wrappers_are_not_loose_equal_to_each_other => {
        r#"
console.log(new String("a") == new String("a"));
"#,
        ["false"]
    };
}