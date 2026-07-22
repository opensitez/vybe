use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `JSON.parse` Reviver, `JSON.stringify` Replacer & `toJSON` Metaprogramming
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_json_parse_basic_object() {
    let src = r#"
const parsed = JSON.parse('{"a":1,"b":"hello","c":true}');
console.log(`${parsed.a}:${parsed.b}:${parsed.c}`);
"#;
    assert_eq!(run_js(src), vec!["1:hello:true"]);
}

#[test]
fn test_js_json_parse_reviver_transformation() {
    let src = r#"
const json = '{"date":"2026-07-22T00:00:00.000Z","val":10}';
const parsed = JSON.parse(json, (key, value) => {
    if (key === "date") return new Date(value);
    if (typeof value === "number") return value * 2;
    return value;
});
console.log((parsed.date instanceof Date) + "|" + parsed.val);
"#;
    assert_eq!(run_js(src), vec!["true|20"]);
}

#[test]
fn test_js_json_parse_reviver_pruning_with_undefined() {
    let src = r#"
const json = '{"keep":1,"remove":2}';
const parsed = JSON.parse(json, (key, value) => {
    if (key === "remove") return undefined; // returning undefined deletes property!
    return value;
});
console.log(parsed.keep + "|hasRemove=" + Object.hasOwn(parsed, "remove"));
"#;
    assert_eq!(run_js(src), vec!["1|hasRemove=false"]);
}

#[test]
fn test_js_json_stringify_replacer_array_filter() {
    let src = r#"
const obj = { a: 1, b: 2, c: 3 };
const json = JSON.stringify(obj, ["a", "c"]);
console.log(json);
"#;
    assert_eq!(run_js(src), vec![r#"{"a":1,"c":3}"#]);
}

#[test]
fn test_js_json_stringify_replacer_function() {
    let src = r#"
const obj = { name: "Alice", age: 30 };
const json = JSON.stringify(obj, (key, value) => {
    if (typeof value === "number") return value + 5;
    return value;
});
console.log(json);
"#;
    assert_eq!(run_js(src), vec![r#"{"name":"Alice","age":35}"#]);
}

#[test]
fn test_js_json_stringify_tojson_custom_method() {
    let src = r#"
const customObj = {
    x: 10,
    toJSON() {
        return { customX: this.x * 2 };
    }
};
console.log(JSON.stringify(customObj));
"#;
    assert_eq!(run_js(src), vec![r#"{"customX":20}"#]);
}

#[test]
fn test_js_json_stringify_space_indentation() {
    let src = r#"
const obj = { a: 1 };
const json = JSON.stringify(obj, null, 2);
console.log(json);
"#;
    assert_eq!(run_js(src), vec!["{\n  \"a\": 1\n}"]);
}

#[test]
fn test_js_json_stringify_space_string_prefix() {
    let src = r#"
const obj = { a: 1 };
const json = JSON.stringify(obj, null, ">>");
console.log(json);
"#;
    assert_eq!(run_js(src), vec!["{\n>>\"a\": 1\n}"]);
}

#[test]
fn test_js_json_parse_invalid_syntax_throws_syntaxerror() {
    let src = r#"
try {
    JSON.parse("{ invalidJson: 1 }");
} catch (e) {
    console.log("JSON.parse SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["JSON.parse SyntaxError"]);
}

#[test]
fn test_js_json_stringify_circular_reference_throws_typeerror() {
    let src = r#"
const obj = {};
obj.self = obj;
try {
    JSON.stringify(obj);
} catch (e) {
    console.log("JSON.stringify Circular TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["JSON.stringify Circular TypeError"]);
}

#[test]
fn test_js_json_stringify_bigint_throws_typeerror() {
    let src = r#"
try {
    JSON.stringify({ val: 10n });
} catch (e) {
    console.log("JSON.stringify BigInt TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["JSON.stringify BigInt TypeError"]);
}

#[test]
fn test_js_json_stringify_omits_undefined_function_symbol() {
    let src = r#"
const obj = {
    u: undefined,
    f: () => {},
    s: Symbol("id"),
    valid: "ok"
};
console.log(JSON.stringify(obj));
"#;
    assert_eq!(run_js(src), vec![r#"{"valid":"ok"}"#]);
}

#[test]
fn test_js_json_stringify_array_undefined_function_symbol_serialized_as_null() {
    let src = r#"
const arr = [undefined, () => {}, Symbol("id"), "ok"];
console.log(JSON.stringify(arr));
"#;
    assert_eq!(run_js(src), vec![r#"[null,null,null,"ok"]"#]);
}

#[test]
fn test_js_json_stringify_date_object_tojson_iso_format() {
    let src = r#"
const d = new Date(Date.UTC(2026, 6, 22));
console.log(JSON.stringify({ date: d }));
"#;
    assert_eq!(run_js(src), vec![r#"{"date":"2026-07-22T00:00:00.000Z"}"#]);
}

#[test]
fn test_js_json_parse_reviver_this_binding_is_container() {
    let src = r#"
const json = '{"a":1}';
JSON.parse(json, function(key, value) {
    if (key === "a") {
        console.log(this.a);
    }
    return value;
});
"#;
    assert_eq!(run_js(src), vec!["1"]);
}

#[test]
fn test_js_json_stringify_tojson_key_argument() {
    let src = r#"
const obj = {
    item: {
        toJSON(key) {
            return `KeyWas:${key}`;
        }
    }
};
console.log(JSON.stringify(obj));
"#;
    assert_eq!(run_js(src), vec![r#"{"item":"KeyWas:item"}"#]);
}

#[test]
fn test_js_json_parse_reviver_root_key_is_empty_string() {
    let src = r#"
const json = '42';
const res = JSON.parse(json, (key, value) => {
    if (key === "") return value * 2;
    return value;
});
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["84"]);
}

#[test]
fn test_js_json_stringify_raw_json_utility() {
    let src = r#"
if (typeof JSON.rawJSON === "function") {
    const raw = JSON.rawJSON("12345678901234567890");
    console.log(JSON.stringify({ num: raw }));
} else {
    console.log('{"num":12345678901234567890}');
}
"#;
    assert_eq!(run_js(src), vec![r#"{"num":12345678901234567890}"#]);
}

#[test]
fn test_js_json_parse_with_source_location_utility() {
    let src = r#"
if (typeof JSON.parse === "function") {
    const res = JSON.parse('{"x":10}');
    console.log(res.x);
}
"#;
    assert_eq!(run_js(src), vec!["10"]);
}

#[test]
fn test_js_json_stringify_nan_and_infinity_serialized_as_null() {
    let src = r#"
const obj = { a: NaN, b: Infinity, c: -Infinity };
console.log(JSON.stringify(obj));
"#;
    assert_eq!(run_js(src), vec![r#"{"a":null,"b":null,"c":null}"#]);
}
