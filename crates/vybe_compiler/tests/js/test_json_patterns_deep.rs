/// JSON patterns — parse reviver, stringify replacer, custom toJSON
use super::helpers::run_js;

#[test]
fn json_stringify_basic_types() {
    assert_eq!(
        run_js(
            r#"
console.log(JSON.stringify(42));
console.log(JSON.stringify("hello"));
console.log(JSON.stringify(true));
console.log(JSON.stringify(null));
"#
        ),
        vec!["42", "\"hello\"", "true", "null"]
    );
}

#[test]
fn json_parse_basic_types() {
    assert_eq!(
        run_js(
            r#"
console.log(JSON.parse("42"));
console.log(JSON.parse('"hello"'));
console.log(JSON.parse("true"));
console.log(JSON.parse("null"));
"#
        ),
        vec!["42", "hello", "true", "null"]
    );
}

#[test]
fn json_stringify_with_space_indent() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1, b: [2, 3] };
const pretty = JSON.stringify(obj, null, 2);
console.log(pretty.includes("\n"));
console.log(pretty.includes("  "));
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn json_parse_reviver_transforms() {
    assert_eq!(
        run_js(
            r#"
const json = '{"name":"Alice","birthDate":"2000-01-01"}';
const parsed = JSON.parse(json, (key, val) => {
    if (key === "birthDate") return new Date(val);
    return val;
});
console.log(parsed.name);
console.log(parsed.birthDate instanceof Date);
"#
        ),
        vec!["Alice", "true"]
    );
}

#[test]
fn json_stringify_replacer_function() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1, b: 2, c: 3, secret: "hidden" };
const result = JSON.stringify(obj, (key, val) => {
    if (key === "secret") return undefined;
    return val;
});
console.log(result.includes("secret"));
console.log(result.includes("a"));
"#
        ),
        vec!["false", "true"]
    );
}

#[test]
fn json_stringify_replacer_array_filter() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1, b: 2, c: 3 };
const result = JSON.stringify(obj, ["a", "c"]);
const parsed = JSON.parse(result);
console.log(parsed.a);
console.log(parsed.b);
console.log(parsed.c);
"#
        ),
        vec!["1", "undefined", "3"]
    );
}

#[test]
fn json_custom_tojson() {
    assert_eq!(
        run_js(
            r#"
class Temperature {
    constructor(celsius) { this.celsius = celsius; }
    toJSON() {
        return { celsius: this.celsius, fahrenheit: this.celsius * 9/5 + 32 };
    }
}
const t = new Temperature(100);
const json = JSON.stringify(t);
const parsed = JSON.parse(json);
console.log(parsed.celsius);
console.log(parsed.fahrenheit);
"#
        ),
        vec!["100", "212"]
    );
}

#[test]
fn json_parse_reviver_all_keys() {
    assert_eq!(
        run_js(
            r#"
const calls = [];
const json = '{"a":1,"b":{"c":2}}';
JSON.parse(json, (key, val) => { calls.push(key); return val; });
// Reviver is called bottom-up: leaf keys first
console.log(calls.includes("a"));
console.log(calls.includes("c"));
console.log(calls[calls.length - 1]); // root "" is last
"#
        ),
        vec!["true", "true", ""]
    );
}

#[test]
fn json_stringify_depth() {
    assert_eq!(
        run_js(
            r#"
const deep = { a: { b: { c: { d: 42 } } } };
const json = JSON.stringify(deep);
const parsed = JSON.parse(json);
console.log(parsed.a.b.c.d);
"#
        ),
        vec!["42"]
    );
}

#[test]
fn json_stringify_array_of_mixed() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, "two", null, true, { x: 3 }];
const json = JSON.stringify(arr);
console.log(json);
"#
        ),
        vec!["[1,\"two\",null,true,{\"x\":3}]"]
    );
}

#[test]
fn json_parse_invalid_throws() {
    assert_eq!(
        run_js(
            r#"
let threw = false;
try { JSON.parse("{invalid}"); } catch (e) { threw = e instanceof SyntaxError; }
console.log(threw);
"#
        ),
        vec!["true"]
    );
}
