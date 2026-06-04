/// Destructuring — complex patterns, defaults, nested, computed
use super::helpers::run_js;

#[test]
fn object_destructuring_nested_defaults() {
    assert_eq!(
        run_js(
            r#"
const config = { server: { host: "localhost" }, timeout: 3000 };
const { server: { host, port = 8080 }, timeout, retries = 3 } = config;
console.log(host);
console.log(port);
console.log(timeout);
console.log(retries);
"#
        ),
        vec!["localhost", "8080", "3000", "3"]
    );
}

#[test]
fn array_destructuring_with_rest() {
    assert_eq!(
        run_js(
            r#"
const [first, second, ...rest] = [1, 2, 3, 4, 5];
console.log(first);
console.log(second);
console.log(rest.join(","));
"#
        ),
        vec!["1", "2", "3,4,5"]
    );
}

#[test]
fn destructuring_in_function_params() {
    assert_eq!(
        run_js(
            r#"
function point({ x = 0, y = 0, z = 0 } = {}) {
    return `${x},${y},${z}`;
}
console.log(point({ x: 1, y: 2 }));
console.log(point({ z: 5 }));
console.log(point());
"#
        ),
        vec!["1,2,0", "0,0,5", "0,0,0"]
    );
}

#[test]
fn destructuring_rename_and_default() {
    assert_eq!(
        run_js(
            r#"
const { name: firstName = "Anonymous", age: years = 0 } = { name: "Alice", age: 30 };
console.log(firstName);
console.log(years);
const { x: a = 10, y: b = 20 } = { x: 5 };
console.log(a);
console.log(b);
"#
        ),
        vec!["Alice", "30", "5", "20"]
    );
}

#[test]
fn computed_property_destructuring() {
    assert_eq!(
        run_js(
            r#"
const key = "name";
const { [key]: value } = { name: "Alice" };
console.log(value);
const prop = "age";
const { [prop]: age = 25 } = {};
console.log(age);
"#
        ),
        vec!["Alice", "25"]
    );
}

#[test]
fn destructuring_for_of_entries() {
    assert_eq!(
        run_js(
            r#"
const map = new Map([["a", 1], ["b", 2], ["c", 3]]);
const results = [];
for (const [key, value] of map) results.push(`${key}=${value}`);
console.log(results.join(","));
"#
        ),
        vec!["a=1,b=2,c=3"]
    );
}

#[test]
fn nested_array_object_mix() {
    assert_eq!(
        run_js(
            r#"
const data = { items: [{ id: 1, tags: ["a", "b"] }, { id: 2, tags: ["c"] }] };
const { items: [{ id: firstId, tags: [firstTag] }, { tags: [secondItemTag] }] } = data;
console.log(firstId);
console.log(firstTag);
console.log(secondItemTag);
"#
        ),
        vec!["1", "a", "c"]
    );
}

#[test]
fn object_rest_spread_partial() {
    assert_eq!(
        run_js(
            r#"
const { a, b, ...rest } = { a: 1, b: 2, c: 3, d: 4 };
console.log(a);
console.log(b);
console.log(Object.keys(rest).sort().join(","));
const merged = { ...rest, e: 5 };
console.log(merged.c);
console.log(merged.e);
"#
        ),
        vec!["1", "2", "c,d", "3", "5"]
    );
}

#[test]
fn swap_via_destructuring() {
    assert_eq!(
        run_js(
            r#"
let x = 1, y = 2;
[x, y] = [y, x];
console.log(x);
console.log(y);
let a = "hello", b = "world";
[a, b] = [b, a];
console.log(a);
console.log(b);
"#
        ),
        vec!["2", "1", "world", "hello"]
    );
}

#[test]
fn destructuring_generator_result() {
    assert_eq!(
        run_js(
            r#"
function* entries(obj) {
    for (const [k, v] of Object.entries(obj)) yield [k, v];
}
const obj = { x: 10, y: 20, z: 30 };
const results = [];
for (const [key, val] of entries(obj)) results.push(key + ":" + val);
console.log(results.join(","));
"#
        ),
        vec!["x:10,y:20,z:30"]
    );
}

#[test]
fn default_value_is_expression() {
    assert_eq!(
        run_js(
            r#"
let counter = 0;
function inc() { return ++counter; }
const { a = inc(), b = inc(), c = 5 } = { b: 99 };
console.log(a);
console.log(b);
console.log(c);
console.log(counter);
"#
        ),
        vec!["1", "99", "5", "1"]
    );
}
