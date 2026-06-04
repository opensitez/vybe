use super::helpers::assert_js;

macro_rules! case {
    ($src:expr, [$($expected:expr),* $(,)?]) => {
        assert_js($src, &[$($expected),*]);
    };
}

#[test]
fn json_parse_preserves_nested_whitespace() {
    case!(
        r#"
const obj = JSON.parse('{\n  "a" : 1 , "b" : [ true , null ] }');
console.log(obj.a);
console.log(obj.b[0]);
console.log(obj.b[1] === null);
"#,
        ["1", "true", "true"]
    );
}

#[test]
fn json_parse_allows_trailing_document_space() {
    case!(
        r#"
const obj = JSON.parse('{"a":1}   \n\t');
console.log(obj.a);
"#,
        ["1"]
    );
}

#[test]
fn json_parse_duplicate_keys_last_wins() {
    case!(
        r#"
const obj = JSON.parse('{"a":1,"a":3}');
console.log(obj.a);
"#,
        ["3"]
    );
}

#[test]
fn json_parse_negative_zero_keeps_sign() {
    case!(
        r#"
const value = JSON.parse('-0');
console.log(1 / value);
"#,
        ["-Infinity"]
    );
}

#[test]
fn json_parse_unicode_escape_in_string() {
    case!(
        r#"
const obj = JSON.parse('{"s":"\u0041\u0042"}');
console.log(obj.s);
"#,
        ["AB"]
    );
}

#[test]
fn json_parse_escaped_quote_sequence() {
    case!(
        r#"
const obj = JSON.parse('{"s":"a\\"b"}');
console.log(obj.s);
"#,
        ["a\"b"]
    );
}

#[test]
fn json_parse_escaped_backslash_sequence() {
    case!(
        r#"
const obj = JSON.parse('{"s":"c\\\\d"}');
console.log(obj.s);
console.log(obj.s.length);
"#,
        [r#"c\d"#, "3"]
    );
}

#[test]
fn json_parse_exponent_number_value() {
    case!(
        r#"
const obj = JSON.parse('{"n":1.25e2}');
console.log(obj.n);
"#,
        ["125"]
    );
}

#[test]
fn json_parse_reviver_transforms_numbers() {
    case!(
        r#"
const obj = JSON.parse('{"a":1,"b":2}', (key, value) => {
    return typeof value === "number" ? value * 10 : value;
});
console.log(obj.a);
console.log(obj.b);
"#,
        ["10", "20"]
    );
}

#[test]
fn json_parse_reviver_can_drop_object_property() {
    case!(
        r#"
const obj = JSON.parse('{"a":1,"b":2}', (key, value) => {
    return key === "b" ? undefined : value;
});
console.log(Object.keys(obj).join(","));
console.log(obj.b);
"#,
        ["a", "undefined"]
    );
}

#[test]
fn json_parse_reviver_can_delete_array_element_into_hole() {
    case!(
        r#"
const arr = JSON.parse('[1,2,3]', (key, value) => {
    return key === "1" ? undefined : value;
});
console.log(1 in arr);
console.log(arr[1]);
console.log(JSON.stringify(arr));
"#,
        ["false", "undefined", "[1,null,3]"]
    );
}

#[test]
fn json_parse_reviver_visits_children_before_parents() {
    case!(
        r#"
const seen = [];
JSON.parse('{"outer":{"inner":1},"arr":[2]}', (key, value) => {
    seen.push(key === "" ? "<root>" : key);
    return value;
});
console.log(seen.join(","));
"#,
        ["inner,outer,0,arr,<root>"]
    );
}

#[test]
fn json_parse_reviver_sees_array_indexes_as_strings() {
    case!(
        r#"
JSON.parse('[10]', (key, value) => {
    if (key !== "") {
        console.log(typeof key + ":" + key);
    }
    return value;
});
"#,
        ["string:0"]
    );
}

#[test]
fn json_parse_reviver_can_replace_root_value() {
    case!(
        r#"
const result = JSON.parse('{"a":1,"b":2}', (key, value) => {
    return key === "" ? { total: value.a + value.b } : value;
});
console.log(result.total);
"#,
        ["3"]
    );
}

#[test]
fn json_parse_reviver_can_replace_primitive_root() {
    case!(
        r#"
const result = JSON.parse('5', (key, value) => {
    return key === "" ? value + 1 : value;
});
console.log(result);
"#,
        ["6"]
    );
}

#[test]
fn json_parse_reviver_can_wrap_nested_object_values() {
    case!(
        r#"
const obj = JSON.parse('{"box":{"value":2}}', (key, value) => {
    return key === "box" ? { wrapped: value.value + 1 } : value;
});
console.log(obj.box.wrapped);
"#,
        ["3"]
    );
}

#[test]
fn json_parse_roundtrip_normalizes_spacing() {
    case!(
        r#"
const compact = JSON.stringify(JSON.parse('{ "a" : [1, 2], "b" : true }'));
console.log(compact);
"#,
        ["{\"a\":[1,2],\"b\":true}"]
    );
}

#[test]
fn json_parse_false_true_null_literals() {
    case!(
        r#"
const obj = JSON.parse('{"t":true,"f":false,"n":null}');
console.log(obj.t);
console.log(obj.f);
console.log(obj.n === null);
"#,
        ["true", "false", "true"]
    );
}

#[test]
fn json_parse_array_root_can_be_reduced_by_reviver() {
    case!(
        r#"
const result = JSON.parse('[1,2,3]', (key, value) => {
    return key === "" ? value.length : value;
});
console.log(result);
"#,
        ["3"]
    );
}

#[test]
fn json_parse_string_escape_roundtrip_matches_input() {
    case!(
        r#"
const text = JSON.parse('"line1\\nline2"');
console.log(text.split("\n").join("|"));
"#,
        ["line1|line2"]
    );
}

#[test]
fn json_stringify_preserves_insertion_order_for_strings() {
    case!(
        r#"
const obj = {};
obj.z = 1;
obj.a = 2;
obj.m = 3;
console.log(JSON.stringify(obj));
"#,
        ["{\"z\":1,\"a\":2,\"m\":3}"]
    );
}

#[test]
fn json_stringify_orders_integer_like_keys_before_strings() {
    case!(
        r#"
const obj = {};
obj.b = 1;
obj["10"] = 2;
obj.a = 3;
obj["2"] = 4;
console.log(JSON.stringify(obj));
"#,
        ["{\"2\":4,\"10\":2,\"b\":1,\"a\":3}"]
    );
}

#[test]
fn json_stringify_sparse_array_holes_become_null() {
    case!(
        r#"
const arr = [];
arr[1] = 2;
arr[3] = 4;
console.log(JSON.stringify(arr));
"#,
        ["[null,2,null,4]"]
    );
}

#[test]
fn json_stringify_array_undefined_and_function_become_null() {
    case!(
        r#"
console.log(JSON.stringify([undefined, function () {}, 3]));
"#,
        ["[null,null,3]"]
    );
}

#[test]
fn json_stringify_object_undefined_and_function_are_omitted() {
    case!(
        r#"
console.log(JSON.stringify({ a: 1, b: undefined, c: function () {}, d: null }));
"#,
        ["{\"a\":1,\"d\":null}"]
    );
}

#[test]
fn json_stringify_top_level_undefined_returns_undefined() {
    case!(
        r#"
console.log(JSON.stringify(undefined) === undefined);
"#,
        ["true"]
    );
}

#[test]
fn json_stringify_top_level_function_returns_undefined() {
    case!(
        r#"
console.log(JSON.stringify(function () {}) === undefined);
"#,
        ["true"]
    );
}

#[test]
fn json_stringify_negative_zero_serializes_as_zero() {
    case!(
        r#"
console.log(JSON.stringify(-0));
"#,
        ["0"]
    );
}

#[test]
fn json_stringify_nan_and_infinity_in_object_become_null() {
    case!(
        r#"
console.log(JSON.stringify({ a: NaN, b: Infinity, c: -Infinity }));
"#,
        ["{\"a\":null,\"b\":null,\"c\":null}"]
    );
}

#[test]
fn json_stringify_nan_and_infinity_in_array_become_null() {
    case!(
        r#"
console.log(JSON.stringify([NaN, Infinity, -Infinity]));
"#,
        ["[null,null,null]"]
    );
}

#[test]
fn json_stringify_inherited_properties_are_ignored() {
    case!(
        r#"
const base = { skip: 1 };
const obj = Object.create(base);
obj.keep = 2;
console.log(JSON.stringify(obj));
"#,
        ["{\"keep\":2}"]
    );
}

#[test]
fn json_stringify_preserves_empty_string_key() {
    case!(
        r#"
console.log(JSON.stringify({ "": 1, a: 2 }));
"#,
        ["{\"\":1,\"a\":2}"]
    );
}

#[test]
fn json_stringify_wrapper_number_object_unboxes() {
    case!(
        r#"
console.log(JSON.stringify(new Number(5)));
"#,
        ["5"]
    );
}

#[test]
fn json_stringify_wrapper_string_object_unboxes() {
    case!(
        r#"
console.log(JSON.stringify(new String("hi")));
"#,
        [r#""hi""#]
    );
}

#[test]
fn json_stringify_wrapper_boolean_object_unboxes() {
    case!(
        r#"
console.log(JSON.stringify(new Boolean(false)));
"#,
        ["false"]
    );
}

#[test]
fn json_stringify_replacer_array_filters_top_level_keys() {
    case!(
        r#"
console.log(JSON.stringify({ a: 1, b: 2, c: 3 }, ["c", "a"]));
"#,
        ["{\"c\":3,\"a\":1}"]
    );
}

#[test]
fn json_stringify_replacer_array_coerces_numeric_entries() {
    case!(
        r#"
console.log(JSON.stringify({ 1: "one", 2: "two", x: "ex" }, [2, "x"]));
"#,
        ["{\"2\":\"two\",\"x\":\"ex\"}"]
    );
}

#[test]
fn json_stringify_replacer_array_applies_to_nested_objects() {
    case!(
        r#"
const obj = { outer: { a: 1, b: 2 }, a: 9 };
console.log(JSON.stringify(obj, ["outer", "b"]));
"#,
        ["{\"outer\":{\"b\":2}}"]
    );
}

#[test]
fn json_stringify_replacer_function_can_scale_numbers() {
    case!(
        r#"
console.log(JSON.stringify({ a: 1, b: 2 }, (key, value) => {
    return typeof value === "number" ? value * 10 : value;
}));
"#,
        ["{\"a\":10,\"b\":20}"]
    );
}

#[test]
fn json_stringify_replacer_function_can_prune_property() {
    case!(
        r#"
console.log(JSON.stringify({ a: 1, b: 2, c: 3 }, (key, value) => {
    return key === "b" ? undefined : value;
}));
"#,
        ["{\"a\":1,\"c\":3}"]
    );
}

#[test]
fn json_stringify_replacer_function_receives_root_empty_key() {
    case!(
        r#"
JSON.stringify({ a: 1 }, (key, value) => {
    if (key === "") {
        console.log("root");
    }
    return value;
});
"#,
        ["root"]
    );
}

#[test]
fn json_stringify_replacer_function_runs_after_tojson() {
    case!(
        r#"
const obj = {
    toJSON() {
        return { x: 2 };
    }
};
console.log(JSON.stringify(obj, (key, value) => {
    return typeof value === "number" ? value * 2 : value;
}));
"#,
        ["{\"x\":4}"]
    );
}

#[test]
fn json_stringify_replacer_function_can_transform_array_elements() {
    case!(
        r#"
console.log(JSON.stringify([1, 2, 3], (key, value) => {
    return typeof value === "number" ? value + 1 : value;
}));
"#,
        ["[2,3,4]"]
    );
}

#[test]
fn json_stringify_replacer_function_sees_object_children_individually() {
    case!(
        r#"
const seen = [];
JSON.stringify({ a: 1, b: { c: 2 } }, (key, value) => {
    if (key !== "") {
        seen.push(key);
    }
    return value;
});
console.log(seen.join(","));
"#,
        ["a,b,c"]
    );
}

#[test]
fn json_stringify_replacer_array_reorders_keys_by_list_order() {
    case!(
        r#"
console.log(JSON.stringify({ a: 1, b: 2, c: 3 }, ["b", "a"]));
"#,
        ["{\"b\":2,\"a\":1}"]
    );
}

#[test]
fn json_stringify_replacer_function_can_replace_nested_object_with_primitive() {
    case!(
        r#"
console.log(JSON.stringify({ a: { b: 2 } }, (key, value) => {
    return key === "a" ? "boxed" : value;
}));
"#,
        ["{\"a\":\"boxed\"}"]
    );
}

#[test]
fn json_stringify_space_number_indents_nested_level() {
    case!(
        r#"
const json = JSON.stringify({ a: 1, b: { c: 2 } }, null, 2);
console.log(json.indexOf('\n  "b"') >= 0);
console.log(json.indexOf('\n    "c"') >= 0);
"#,
        ["true", "true"]
    );
}

#[test]
fn json_stringify_space_number_is_capped_at_ten() {
    case!(
        r#"
const json = JSON.stringify({ a: 1 }, null, 20);
console.log(json.indexOf('\n          "a"') >= 0);
console.log(json.indexOf('\n           "a"') === -1);
"#,
        ["true", "true"]
    );
}

#[test]
fn json_stringify_space_string_uses_first_ten_chars() {
    case!(
        r#"
const json = JSON.stringify({ a: { b: 1 } }, null, "abcdefghijklm");
console.log(json.indexOf('\nabcdefghijabcdefghij"b"') >= 0);
"#,
        ["true"]
    );
}

#[test]
fn json_stringify_zero_space_stays_compact() {
    case!(
        r#"
console.log(JSON.stringify({ a: 1, b: 2 }, null, 0));
"#,
        ["{\"a\":1,\"b\":2}"]
    );
}

#[test]
fn json_stringify_escapes_newline_and_tab() {
    case!(
        r#"
const json = JSON.stringify({ s: "a\n\tb" });
console.log(json.indexOf("\\n") >= 0);
console.log(json.indexOf("\\t") >= 0);
"#,
        ["true", "true"]
    );
}

#[test]
fn json_stringify_escapes_quote_and_backslash() {
    case!(
        r#"
const json = JSON.stringify({ s: 'a"\\b' });
console.log(json.indexOf('\\"') >= 0);
console.log(json.indexOf('\\\\') >= 0);
"#,
        ["true", "true"]
    );
}

#[test]
fn json_stringify_escapes_backspace_and_formfeed() {
    case!(
        r#"
const json = JSON.stringify({ s: "a\b\fb" });
console.log(json.indexOf("\\b") >= 0);
console.log(json.indexOf("\\f") >= 0);
"#,
        ["true", "true"]
    );
}

#[test]
fn json_stringify_compact_then_pretty_outputs_same_data() {
    case!(
        r#"
const src = { a: 1, b: [2, 3] };
const a = JSON.parse(JSON.stringify(src));
const b = JSON.parse(JSON.stringify(src, null, 2));
console.log(a.b[1] === b.b[1]);
console.log(JSON.stringify(a) === JSON.stringify(b));
"#,
        ["true", "true"]
    );
}

#[test]
fn json_stringify_tojson_on_root_object_controls_output() {
    case!(
        r#"
const obj = {
    a: 1,
    toJSON() {
        return { b: 2 };
    }
};
console.log(JSON.stringify(obj));
"#,
        ["{\"b\":2}"]
    );
}

#[test]
fn json_stringify_tojson_on_nested_object_controls_output() {
    case!(
        r#"
const obj = {
    wrap: {
        a: 1,
        toJSON() {
            return "x";
        }
    }
};
console.log(JSON.stringify(obj));
"#,
        ["{\"wrap\":\"x\"}"]
    );
}

#[test]
fn json_stringify_tojson_can_return_primitive() {
    case!(
        r#"
const obj = {
    toJSON() {
        return 5;
    }
};
console.log(JSON.stringify(obj));
"#,
        ["5"]
    );
}

#[test]
fn json_stringify_tojson_receives_own_key_for_nested_property() {
    case!(
        r#"
const outer = {
    inner: {
        toJSON(key) {
            console.log(key);
            return 1;
        }
    }
};
console.log(JSON.stringify(outer));
"#,
        ["inner", "{\"inner\":1}"]
    );
}

#[test]
fn json_stringify_tojson_precedes_replacer_observation() {
    case!(
        r#"
const obj = {
    inner: {
        toJSON() {
            return { x: 1 };
        }
    }
};
JSON.stringify(obj, (key, value) => {
    if (key === "inner") {
        console.log(value.x);
    }
    return value;
});
"#,
        ["1"]
    );
}

#[test]
fn json_stringify_array_element_tojson_controls_slot_value() {
    case!(
        r#"
const arr = [{
    toJSON() {
        return "v";
    }
}];
console.log(JSON.stringify(arr));
"#,
        ["[\"v\"]"]
    );
}

#[test]
fn json_roundtrip_preserves_null_false_zero_and_empty_string() {
    case!(
        r#"
const src = { n: null, f: false, z: 0, s: "" };
const back = JSON.parse(JSON.stringify(src));
console.log(back.n === null);
console.log(back.f);
console.log(back.z);
console.log(back.s === "");
"#,
        ["true", "false", "0", "true"]
    );
}

#[test]
fn json_roundtrip_drops_object_undefined_fields() {
    case!(
        r#"
const back = JSON.parse(JSON.stringify({ a: 1, b: undefined }));
console.log(Object.keys(back).join(","));
console.log(back.b);
"#,
        ["a", "undefined"]
    );
}

#[test]
fn json_roundtrip_keeps_array_null_placeholders() {
    case!(
        r#"
const back = JSON.parse(JSON.stringify([1, null, 3]));
console.log(back[1] === null);
"#,
        ["true"]
    );
}

#[test]
fn json_roundtrip_preserves_nested_arrays_and_objects() {
    case!(
        r#"
const src = { a: [{ b: 2 }, 3] };
const back = JSON.parse(JSON.stringify(src));
console.log(back.a[0].b);
console.log(back.a[1]);
"#,
        ["2", "3"]
    );
}

#[test]
fn json_roundtrip_preserves_escaped_control_text() {
    case!(
        r#"
const src = { s: "line1\nline2\tend" };
const back = JSON.parse(JSON.stringify(src));
console.log(back.s === src.s);
"#,
        ["true"]
    );
}

#[test]
fn json_roundtrip_with_replacer_and_reviver_can_recover_numbers() {
    case!(
        r#"
const src = { a: 2, b: 4 };
const json = JSON.stringify(src, (key, value) => {
    return typeof value === "number" ? value * 2 : value;
});
const back = JSON.parse(json, (key, value) => {
    return typeof value === "number" ? value / 2 : value;
});
console.log(back.a);
console.log(back.b);
"#,
        ["2", "4"]
    );
}

#[test]
fn json_stringify_then_parse_preserves_property_order() {
    case!(
        r#"
const src = {};
src.c = 1;
src.a = 2;
const back = JSON.parse(JSON.stringify(src));
console.log(Object.keys(back).join(","));
"#,
        ["c,a"]
    );
}

#[test]
fn json_parse_then_stringify_sorts_integer_like_keys() {
    case!(
        r#"
const json = JSON.stringify(JSON.parse('{"10":1,"2":2,"a":3}'));
console.log(json);
"#,
        ["{\"2\":2,\"10\":1,\"a\":3}"]
    );
}

#[test]
fn json_parse_reviver_can_turn_strings_into_numbers() {
    case!(
        r#"
const obj = JSON.parse('{"a":"2","b":"3"}', (key, value) => {
    return typeof value === "string" ? Number(value) : value;
});
console.log(obj.a + obj.b);
"#,
        ["5"]
    );
}

#[test]
fn json_stringify_custom_tojson_and_parse_reviver_compose() {
    case!(
        r#"
const src = {
    item: {
        value: 4,
        toJSON() {
            return { value: this.value * 2 };
        }
    }
};
const back = JSON.parse(JSON.stringify(src), (key, value) => {
    return key === "value" ? value / 2 : value;
});
console.log(back.item.value);
"#,
        ["4"]
    );
}
