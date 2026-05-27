crate::js_cases! {
    number_empty_array_is_zero => { r#"console.log(Number([]));"#, ["0"] };
    number_singleton_array_is_element_number => { r#"console.log(Number([7]));"#, ["7"] };
    number_multi_element_array_is_nan => { r#"console.log(Number.isNaN(Number([1,2])));"#, ["true"] };
    number_object_valueof_is_used => {
        r#"
const value = Number({ valueOf() { return 7; } });
console.log(value);
"#,
        ["7"]
    };
    number_object_tostring_is_used_when_valueof_not_primitive => {
        r#"
const value = Number({ valueOf() { return {}; }, toString() { return "8"; } });
console.log(value);
"#,
        ["8"]
    };
    string_empty_array_is_empty_string => { r#"console.log(String([]) === "");"#, ["true"] };
    string_singleton_array_is_element_text => { r#"console.log(String([7]));"#, ["7"] };
    string_symbol_for_uses_registry_symbol_format => { r#"console.log(String(Symbol.for("reg")));"#, ["Symbol(reg)"] };
    boolean_symbol_is_true => { r#"console.log(Boolean(Symbol("x")));"#, ["true"] };
    boolean_bigint_zero_is_false => { r#"console.log(Boolean(0n));"#, ["false"] };
    boolean_bigint_nonzero_is_true => { r#"console.log(Boolean(1n));"#, ["true"] };
    number_bigint_throws_typeerror => {
        r#"
try {
  Number(1n);
  console.log("no error");
} catch (error) {
  console.log(error instanceof TypeError);
}
"#,
        ["true"]
    };
    object_boolean_primitive_boxes_to_object => { r#"console.log(typeof Object(true));"#, ["object"] };
    object_number_primitive_boxes_to_object => { r#"console.log(typeof Object(1));"#, ["object"] };
    object_string_primitive_boxes_to_object => { r#"console.log(typeof Object("x"));"#, ["object"] };
    object_symbol_primitive_boxes_to_object => { r#"console.log(typeof Object(Symbol("x")));"#, ["object"] };
    object_bigint_primitive_boxes_to_object => { r#"console.log(typeof Object(1n));"#, ["object"] };
    string_function_uses_source_text_prefix => {
        r#"
const text = String(function demo(a, b) { return a + b; });
console.log(text.startsWith("function demo"));
"#,
        ["true"]
    };
    boolean_function_object_is_true => { r#"console.log(Boolean(function() {}));"#, ["true"] };
    number_true_string_with_whitespace_parses => { r#"console.log(Number("\n42\t"));"#, ["42"] };
    number_plus_string_parses => { r#"console.log(Number("+42"));"#, ["42"] };
    number_minus_string_parses => { r#"console.log(Number("-42"));"#, ["-42"] };
    string_negative_zero_is_plain_zero => { r#"console.log(String(-0));"#, ["0"] };
    string_infinity_is_literal_infinity => { r#"console.log(String(Infinity));"#, ["Infinity"] };
    string_negative_infinity_is_literal_negative_infinity => { r#"console.log(String(-Infinity));"#, ["-Infinity"] };
    number_false_string_is_nan => { r#"console.log(Number.isNaN(Number("false")));"#, ["true"] };
    boolean_date_object_is_true => { r#"console.log(Boolean(new Date(0)));"#, ["true"] };
    string_date_object_uses_date_string_representation => {
        r#"
console.log(String(new Date(0)).length > 0);
"#,
        ["true"]
    };
    object_preserves_existing_object_identity => {
        r#"
const obj = {};
console.log(Object(obj) === obj);
"#,
        ["true"]
    };
    boolean_object_wrapper_is_true => { r#"console.log(Boolean(Object(false)));"#, ["true"] };
    number_infinity_string_parses_to_infinity => { r#"console.log(Number("Infinity"));"#, ["Infinity"] };
    number_negative_infinity_string_parses => { r#"console.log(Number("-Infinity"));"#, ["-Infinity"] };
    number_positive_zero_string_parses => { r#"console.log(Number("+0"));"#, ["0"] };
    number_negative_zero_string_preserves_sign => { r#"console.log(1 / Number("-0"));"#, ["-Infinity"] };
    number_scientific_notation_string_parses => { r#"console.log(Number("1e3"));"#, ["1000"] };
    number_numeric_separator_string_is_nan => { r#"console.log(Number.isNaN(Number("1_000")));"#, ["true"] };
    number_leading_zero_decimal_string_parses => { r#"console.log(Number("08"));"#, ["8"] };
    number_invalid_binary_literal_string_is_nan => { r#"console.log(Number.isNaN(Number("0b2")));"#, ["true"] };
    number_hex_string_with_whitespace_parses => { r#"console.log(Number(" 0xF "));"#, ["15"] };
    boolean_false_word_string_is_true => { r#"console.log(Boolean("false"));"#, ["true"] };
    boolean_space_string_is_true => { r#"console.log(Boolean(" "));"#, ["true"] };
    string_bigint_is_decimal_text => { r#"console.log(String(12n));"#, ["12"] };
    string_symbol_without_description_uses_empty_symbol_format => { r#"console.log(String(Symbol()));"#, ["Symbol()"] };
    object_number_box_valueof_roundtrips_primitive => { r#"console.log(Object(1).valueOf());"#, ["1"] };
    object_boolean_box_valueof_roundtrips_primitive => { r#"console.log(Object(true).valueOf());"#, ["true"] };
    object_string_box_valueof_roundtrips_primitive => { r#"console.log(Object("x").valueOf());"#, ["x"] };
    object_bigint_box_valueof_roundtrips_primitive => { r#"console.log(Object(12n).valueOf());"#, ["12n"] };
    object_symbol_box_valueof_roundtrips_primitive => {
        r#"
const s = Symbol.for("boxed");
console.log(Object(s).valueOf() === s);
"#,
        ["true"]
    };
    number_object_with_string_valueof_coerces_string_primitive => {
        r#"
console.log(Number({ valueOf() { return "11"; } }));
"#,
        ["11"]
    };
    number_object_with_only_tostring_can_still_coerce => {
        r#"
console.log(Number({ toString() { return " 9 "; } }));
"#,
        ["9"]
    };
    string_object_prefers_tostring_for_string_hint => {
        r#"
console.log(String({ toString() { return "ok"; }, valueOf() { return 3; } }));
"#,
        ["ok"]
    };
    string_array_with_null_and_undefined_joins_empty_slots => { r#"console.log(String([null, undefined]));"#, [","] };
    string_array_with_single_null_is_empty_string => { r#"console.log(String([null]) === "");"#, ["true"] };
    number_single_true_array_is_nan => { r#"console.log(Number.isNaN(Number([true])));"#, ["true"] };
    number_single_null_array_is_zero => { r#"console.log(Number([null]));"#, ["0"] };
    object_null_constructor_is_object => { r#"console.log(Object(null).constructor === Object);"#, ["true"] };
    object_undefined_constructor_is_object => { r#"console.log(Object(undefined).constructor === Object);"#, ["true"] };
    boolean_boxed_number_zero_is_true => { r#"console.log(Boolean(Object(0)));"#, ["true"] };
    string_empty_object_has_object_prototype_tag => { r#"console.log(String(Object.create({})));"#, ["[object Object]"] };
    number_boolean_object_is_one_or_zero_via_valueof => {
        r#"
console.log(Number(Object(true).valueOf()));
"#,
        ["1"]
    };
}