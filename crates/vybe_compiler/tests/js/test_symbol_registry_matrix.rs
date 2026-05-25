crate::js_cases! {
    symbol_for_same_key_returns_same_symbol => {
        r#"
console.log(Symbol.for("shared") === Symbol.for("shared"));
"#,
        ["true"]
    };

    symbol_for_different_keys_returns_distinct_symbols => {
        r#"
console.log(Symbol.for("a") === Symbol.for("b"));
"#,
        ["false"]
    };

    symbol_keyfor_returns_registered_key => {
        r#"
console.log(Symbol.keyFor(Symbol.for("shared")));
"#,
        ["shared"]
    };

    symbol_keyfor_for_local_symbol_is_undefined => {
        r#"
console.log(Symbol.keyFor(Symbol("local")) === undefined);
"#,
        ["true"]
    };

    symbol_description_is_preserved => {
        r#"
console.log(Symbol("demo").description);
"#,
        ["demo"]
    };

    symbol_without_description_has_undefined_description => {
        r#"
console.log(Symbol().description === undefined);
"#,
        ["true"]
    };

    symbol_empty_string_description_is_empty => {
        r#"
console.log(Symbol("").description === "");
"#,
        ["true"]
    };

    symbol_tostring_includes_description => {
        r#"
console.log(Symbol("demo").toString());
"#,
        ["Symbol(demo)"]
    };

    distinct_local_symbols_are_not_equal => {
        r#"
console.log(Symbol("x") === Symbol("x"));
"#,
        ["false"]
    };

    string_conversion_of_symbol_uses_symbol_format => {
        r#"
console.log(String(Symbol("x")));
"#,
        ["Symbol(x)"]
    };

    symbol_can_be_used_as_object_key => {
        r#"
const s = Symbol("id");
const obj = { [s]: 42 };
console.log(obj[s]);
"#,
        ["42"]
    };

    symbol_keys_are_skipped_by_object_keys => {
        r#"
const s = Symbol("id");
const obj = { visible: 1, [s]: 2 };
console.log(Object.keys(obj).join(","));
"#,
        ["visible"]
    };

    symbol_keys_appear_in_getownpropertysymbols => {
        r#"
const a = Symbol("a");
const b = Symbol("b");
const obj = { [a]: 1, [b]: 2 };
console.log(Object.getOwnPropertySymbols(obj).length);
"#,
        ["2"]
    };

    symbol_property_can_be_enumerable_even_if_object_keys_skips_it => {
        r#"
const s = Symbol("a");
const obj = {};
obj[s] = 1;
console.log(obj.propertyIsEnumerable(s));
"#,
        ["true"]
    };

    json_stringify_ignores_symbol_keyed_properties => {
        r#"
const s = Symbol("a");
const obj = { visible: 1, [s]: 2 };
console.log(JSON.stringify(obj));
"#,
        ["{\"visible\":1}"]
    };

    json_stringify_symbol_value_in_object_omits_property => {
        r#"
console.log(JSON.stringify({ x: Symbol("a") }));
"#,
        ["{}"]
    };

    json_stringify_symbol_in_array_becomes_null => {
        r#"
console.log(JSON.stringify([Symbol("a")]));
"#,
        ["[null]"]
    };

    object_getownpropertysymbols_preserves_insertion_order => {
        r#"
const a = Symbol("a");
const b = Symbol("b");
const obj = {};
obj[a] = 1;
obj[b] = 2;
const syms = Object.getOwnPropertySymbols(obj);
console.log(syms[0] === a);
console.log(syms[1] === b);
"#,
        ["true", "true"]
    };

    symbol_for_empty_key_roundtrips_through_keyfor => {
        r#"
const s = Symbol.for("");
console.log(Symbol.keyFor(s) === "");
"#,
        ["true"]
    };

    symbol_registry_value_can_index_object_consistently => {
        r#"
const s1 = Symbol.for("reg");
const s2 = Symbol.for("reg");
const obj = { [s1]: 99 };
console.log(obj[s2]);
"#,
        ["99"]
    };

    implicit_symbol_concatenation_throws_typeerror => {
        r#"
try {
  console.log("x" + Symbol("a"));
} catch (error) {
  console.log(error instanceof TypeError);
}
"#,
        ["true"]
    };

    symbol_for_key_is_stringified => {
        r#"
console.log(Symbol.keyFor(Symbol.for(42)));
"#,
        ["42"]
    };

    object_prototype_tostring_of_symbol_primitive_reports_symbol => {
        r#"
console.log(Object.prototype.toString.call(Symbol("a")));
"#,
        ["[object Symbol]"]
    };

    symbol_valueof_returns_same_symbol => {
        r#"
const s = Symbol("a");
console.log(Object(s).valueOf() === s);
"#,
        ["true"]
    };
}