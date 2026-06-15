/// JSON and serialization patterns
use super::helpers::run_js;

#[test]
fn json_roundtrip_types() {
    assert_eq!(
        run_js(
            r#"
const original = { num: 42, str: "hello", bool: true, arr: [1,2,3], nil: null };
const json = JSON.stringify(original);
const parsed = JSON.parse(json);
console.log(parsed.num);
console.log(parsed.str);
console.log(parsed.bool);
console.log(parsed.arr.join(","));
console.log(parsed.nil);
"#
        ),
        vec!["42", "hello", "true", "1,2,3", "null"]
    );
}

#[test]
fn json_custom_replacer_function() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1, b: undefined, c: null, d: function(){}, e: 2 };
const json = JSON.stringify(obj, (key, val) => {
    if (val === undefined || typeof val === "function") return undefined;
    return val;
});
const parsed = JSON.parse(json);
console.log(parsed.a);
console.log("b" in parsed);
console.log("d" in parsed);
console.log(parsed.e);
"#
        ),
        vec!["1", "false", "false", "2"]
    );
}

#[test]
fn json_replacer_array_filter() {
    assert_eq!(
        run_js(
            r#"
const obj = { name: "Alice", age: 30, password: "secret", email: "a@b.com" };
const json = JSON.stringify(obj, ["name", "email"]);
const parsed = JSON.parse(json);
console.log(parsed.name);
console.log(parsed.email);
console.log("age" in parsed);
console.log("password" in parsed);
"#
        ),
        vec!["Alice", "a@b.com", "false", "false"]
    );
}

#[test]
fn json_space_formatting() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1, b: [2, 3] };
const pretty = JSON.stringify(obj, null, 2);
const lines = pretty.split("\n");
console.log(lines.length > 1);
console.log(lines[0]);
"#
        ),
        vec!["true", "{"]
    );
}

#[test]
#[allow(non_snake_case)]
fn json_custom_toJSON() {
    assert_eq!(
        run_js(
            r#"
class Money {
    constructor(amount, currency) { this.amount = amount; this.currency = currency; }
    toJSON() { return `${this.currency}${this.amount}`; }
}
const obj = { price: new Money(100, "$"), tax: new Money(10, "$") };
const json = JSON.parse(JSON.stringify(obj));
console.log(json.price);
console.log(json.tax);
"#
        ),
        vec!["$100", "$10"]
    );
}

#[test]
fn json_reviver_transform() {
    assert_eq!(
        run_js(
            r#"
const json = '{"createdAt":"2024-01-15T10:00:00.000Z","amount":"42.5"}';
const parsed = JSON.parse(json, (key, val) => {
    if (key === "createdAt") return new Date(val).getFullYear();
    if (key === "amount") return parseFloat(val);
    return val;
});
console.log(parsed.createdAt);
console.log(parsed.amount);
console.log(typeof parsed.amount);
"#
        ),
        vec!["2024", "42.5", "number"]
    );
}

#[test]
fn json_deep_clone() {
    assert_eq!(
        run_js(
            r#"
function deepClone(obj) { return JSON.parse(JSON.stringify(obj)); }
const orig = { a: { b: { c: [1, 2, 3] } } };
const clone = deepClone(orig);
clone.a.b.c.push(4);
console.log(orig.a.b.c.length);
console.log(clone.a.b.c.length);
console.log(orig.a === clone.a);
"#
        ),
        vec!["3", "4", "false"]
    );
}

#[test]
fn json_stringify_undefined_handling() {
    assert_eq!(
        run_js(
            r#"
// JSON.stringify removes undefined values in objects
console.log(JSON.stringify({ a: undefined }));
// but keeps undefined in arrays as null
console.log(JSON.stringify([undefined, 1, undefined]));
// standalone undefined returns undefined (not a string)
console.log(JSON.stringify(undefined));
"#
        ),
        vec!["{}", "[null,1,null]", "undefined"]
    );
}

#[test]
fn json_parse_invalid_throws() {
    assert_eq!(
        run_js(
            r#"
const invalids = ["undefined", "NaN", "{a:1}", "{'a':1}", "[1,2,]"];
let count = 0;
for (const s of invalids) {
    try { JSON.parse(s); }
    catch { count++; }
}
console.log(count);
"#
        ),
        vec!["5"]
    );
}

#[test]
fn json_nested_array_object() {
    assert_eq!(
        run_js(
            r#"
const data = JSON.parse('[{"id":1,"tags":["a","b"]},{"id":2,"tags":["c"]}]');
console.log(data.length);
console.log(data[0].tags.join(","));
console.log(data[1].id);
"#
        ),
        vec!["2", "a,b", "2"]
    );
}
