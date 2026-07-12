/// JSON.parse and JSON.stringify — structured clone, replacer/reviver edge cases,
/// circular reference detection, toJSON override, BigInt error, special values,
/// formatting, null/undefined/function omission.
use super::helpers::run_js;

// ── basic parse/stringify ─────────────────────────────────────────────────────

#[test]
fn stringify_and_parse_roundtrip() {
    assert_eq!(
        run_js(
            r#"
const obj = { name: "Alice", age: 30, active: true };
const json = JSON.stringify(obj);
const parsed = JSON.parse(json);
console.log(parsed.name);
console.log(parsed.age);
console.log(parsed.active);
"#
        ),
        vec!["Alice", "30", "true"]
    );
}

#[test]
fn stringify_null_and_primitives() {
    assert_eq!(
        run_js(
            r#"
console.log(JSON.stringify(null));
console.log(JSON.stringify(42));
console.log(JSON.stringify("hello"));
console.log(JSON.stringify(true));
"#
        ),
        vec!["null", "42", "\"hello\"", "true"]
    );
}

// ── special value omission ────────────────────────────────────────────────────

#[test]
fn stringify_omits_undefined_functions_symbols() {
    assert_eq!(
        run_js(
            r#"
const obj = {
    a: 1,
    b: undefined,
    c: () => {},
    d: Symbol("x"),
    e: "keep"
};
const result = JSON.parse(JSON.stringify(obj));
console.log(result.a);
console.log("b" in result);
console.log("c" in result);
console.log("d" in result);
console.log(result.e);
"#
        ),
        vec!["1", "false", "false", "false", "keep"]
    );
}

#[test]
fn stringify_nan_infinity_become_null() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: NaN, b: Infinity, c: -Infinity };
const result = JSON.parse(JSON.stringify(obj));
console.log(result.a);
console.log(result.b);
console.log(result.c);
"#
        ),
        vec!["null", "null", "null"]
    );
}

#[test]
fn stringify_array_with_undefined_becomes_null() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, undefined, null, 3];
console.log(JSON.stringify(arr));
"#
        ),
        vec!["[1,null,null,3]"]
    );
}

// ── replacer function ─────────────────────────────────────────────────────────

#[test]
fn stringify_replacer_filters_keys() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1, b: 2, c: 3 };
const json = JSON.stringify(obj, (key, value) => key === "b" ? undefined : value);
const result = JSON.parse(json);
console.log(result.a);
console.log("b" in result);
"#
        ),
        vec!["1", "false"]
    );
}

#[test]
fn stringify_replacer_transforms_values() {
    assert_eq!(
        run_js(
            r#"
const data = { score: 0.123456789 };
const json = JSON.stringify(data, (key, value) =>
    typeof value === "number" ? Math.round(value * 100) / 100 : value
);
console.log(json);
"#
        ),
        vec!["{\"score\":0.12}"]
    );
}

#[test]
fn stringify_replacer_array_allows_only_listed_keys() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1, b: 2, c: 3, d: 4 };
const json = JSON.stringify(obj, ["a", "c"]);
const result = JSON.parse(json);
console.log(result.a);
console.log("b" in result);
console.log(result.c);
"#
        ),
        vec!["1", "false", "3"]
    );
}

// ── reviver function ──────────────────────────────────────────────────────────

#[test]
fn parse_reviver_converts_date_strings() {
    assert_eq!(
        run_js(
            r#"
const json = '{"created":"2024-01-15","value":42}';
const result = JSON.parse(json, (key, value) => {
    if (key === "created") return new Date(value).getFullYear();
    return value;
});
console.log(result.created);
console.log(result.value);
"#
        ),
        vec!["2024", "42"]
    );
}

#[test]
fn parse_reviver_receives_all_key_value_pairs() {
    assert_eq!(
        run_js(
            r#"
const keys = [];
JSON.parse('{"a":1,"b":{"c":2}}', (key, value) => {
    if (key !== "") keys.push(key);
    return value;
});
console.log(keys.sort().join(","));
"#
        ),
        vec!["a,b,c"]
    );
}

// ── toJSON ────────────────────────────────────────────────────────────────────

#[test]
fn to_json_method_called_during_stringify() {
    assert_eq!(
        run_js(
            r#"
const obj = {
    value: 42,
    secret: "hidden",
    toJSON() { return { value: this.value }; }
};
const result = JSON.parse(JSON.stringify(obj));
console.log(result.value);
console.log("secret" in result);
"#
        ),
        vec!["42", "false"]
    );
}

#[test]
fn date_to_json_returns_iso_string() {
    assert_eq!(
        run_js(
            r#"
const d = new Date(0); // epoch
const json = JSON.stringify({ date: d });
console.log(json.includes("1970"));
"#
        ),
        vec!["true"]
    );
}

// ── space formatting ──────────────────────────────────────────────────────────

#[test]
fn stringify_with_space_indentation() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1 };
const pretty = JSON.stringify(obj, null, 2);
console.log(pretty.includes("\n"));
console.log(pretty.includes("  "));
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn stringify_with_tab_character() {
    assert_eq!(
        run_js(
            r#"
const obj = { x: 1 };
const result = JSON.stringify(obj, null, "\t");
console.log(result.includes("\t"));
"#
        ),
        vec!["true"]
    );
}

// ── nested and deep structures ────────────────────────────────────────────────

#[test]
fn stringify_deeply_nested() {
    assert_eq!(
        run_js(
            r#"
const deep = { a: { b: { c: { d: 42 } } } };
const result = JSON.parse(JSON.stringify(deep));
console.log(result.a.b.c.d);
"#
        ),
        vec!["42"]
    );
}

#[test]
fn stringify_array_of_objects() {
    assert_eq!(
        run_js(
            r#"
const arr = [{ id: 1, name: "a" }, { id: 2, name: "b" }];
const result = JSON.parse(JSON.stringify(arr));
console.log(result.length);
console.log(result[1].name);
"#
        ),
        vec!["2", "b"]
    );
}

// ── error cases ───────────────────────────────────────────────────────────────

#[test]
fn parse_invalid_json_throws_syntax_error() {
    assert_eq!(
        run_js(
            r#"
let threw = false;
try { JSON.parse("{bad json}"); } catch (e) { threw = e instanceof SyntaxError; }
console.log(threw);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn stringify_circular_reference_throws() {
    assert_eq!(
        run_js(
            r#"
const obj = {};
obj.self = obj;
let threw = false;
try { JSON.stringify(obj); } catch (e) { threw = e instanceof TypeError; }
console.log(threw);
"#
        ),
        vec!["true"]
    );
}
