use super::helpers::run_js;

// ── Array destructuring ───────────────────────────────────
#[test]
fn array_destructure_basic() {
    assert_eq!(
        run_js(
            r#"
const [a, b, c] = [1, 2, 3];
console.log(a, b, c);
"#
        ),
        vec!["1 2 3"]
    );
}

#[test]
fn array_destructure_skip_elements() {
    assert_eq!(
        run_js(
            r#"
const [,, third] = [10, 20, 30];
console.log(third);
"#
        ),
        vec!["30"]
    );
}

#[test]
fn array_destructure_rest() {
    assert_eq!(
        run_js(
            r#"
const [first, ...rest] = [1, 2, 3, 4];
console.log(first);
console.log(rest.join(","));
"#
        ),
        vec!["1", "2,3,4"]
    );
}

#[test]
fn array_destructure_default_values() {
    assert_eq!(
        run_js(
            r#"
const [a = 10, b = 20] = [1];
console.log(a);
console.log(b);
"#
        ),
        vec!["1", "20"]
    );
}

#[test]
fn array_destructure_swap_variables() {
    assert_eq!(
        run_js(
            r#"
let x = 1, y = 2;
[x, y] = [y, x];
console.log(x, y);
"#
        ),
        vec!["2 1"]
    );
}

#[test]
fn array_destructure_nested() {
    assert_eq!(
        run_js(
            r#"
const [[a, b], [c, d]] = [[1, 2], [3, 4]];
console.log(a + b + c + d);
"#
        ),
        vec!["10"]
    );
}

#[test]
fn array_destructure_from_function() {
    assert_eq!(
        run_js(
            r#"
function getCoords() { return [10, 20]; }
const [x, y] = getCoords();
console.log(x + y);
"#
        ),
        vec!["30"]
    );
}

#[test]
fn array_destructure_string_iterating() {
    assert_eq!(
        run_js(
            r#"
const [a, b, c] = "xyz";
console.log(a, b, c);
"#
        ),
        vec!["x y z"]
    );
}

// ── Object destructuring ──────────────────────────────────
#[test]
fn object_destructure_basic() {
    assert_eq!(
        run_js(
            r#"
const { name, age } = { name: "Alice", age: 30 };
console.log(name, age);
"#
        ),
        vec!["Alice 30"]
    );
}

#[test]
fn object_destructure_rename() {
    assert_eq!(
        run_js(
            r#"
const { name: fullName, age: years } = { name: "Bob", age: 25 };
console.log(fullName, years);
"#
        ),
        vec!["Bob 25"]
    );
}

#[test]
fn object_destructure_default_values() {
    assert_eq!(
        run_js(
            r#"
const { a = 1, b = 2 } = { a: 10 };
console.log(a, b);
"#
        ),
        vec!["10 2"]
    );
}

#[test]
fn object_destructure_rest() {
    assert_eq!(
        run_js(
            r#"
const { x, ...rest } = { x: 1, y: 2, z: 3 };
console.log(x);
console.log(Object.keys(rest).sort().join(","));
"#
        ),
        vec!["1", "y,z"]
    );
}

#[test]
fn object_destructure_nested() {
    assert_eq!(
        run_js(
            r#"
const { a: { b: { c } } } = { a: { b: { c: 42 } } };
console.log(c);
"#
        ),
        vec!["42"]
    );
}

#[test]
fn object_destructure_computed_key() {
    assert_eq!(
        run_js(
            r#"
const key = "name";
const { [key]: value } = { name: "Carol" };
console.log(value);
"#
        ),
        vec!["Carol"]
    );
}

#[test]
fn object_destructure_rename_with_default() {
    assert_eq!(
        run_js(
            r#"
const { x: px = 0, y: py = 0 } = { x: 5 };
console.log(px, py);
"#
        ),
        vec!["5 0"]
    );
}

#[test]
fn object_destructure_in_function_params() {
    assert_eq!(
        run_js(
            r#"
function greet({ name, greeting = "Hello" }) {
  console.log(greeting + ", " + name + "!");
}
greet({ name: "Dave" });
greet({ name: "Eve", greeting: "Hi" });
"#
        ),
        vec!["Hello, Dave!", "Hi, Eve!"]
    );
}

#[test]
fn object_destructure_from_class_instance() {
    assert_eq!(
        run_js(
            r#"
class Point {
  constructor(x, y) { this.x = x; this.y = y; }
}
const { x, y } = new Point(3, 4);
console.log(x + y);
"#
        ),
        vec!["7"]
    );
}

// ── Mixed destructuring ───────────────────────────────────
#[test]
fn mixed_destructure_array_of_objects() {
    assert_eq!(
        run_js(
            r#"
const [{ name: n1 }, { name: n2 }] = [{ name: "A" }, { name: "B" }];
console.log(n1, n2);
"#
        ),
        vec!["A B"]
    );
}

#[test]
fn mixed_destructure_object_with_arrays() {
    assert_eq!(
        run_js(
            r#"
const { nums: [a, b], tag } = { nums: [1, 2], tag: "ok" };
console.log(a, b, tag);
"#
        ),
        vec!["1 2 ok"]
    );
}

#[test]
fn destructure_in_for_of() {
    assert_eq!(
        run_js(
            r#"
const pairs = [["a", 1], ["b", 2]];
const keys = [];
for (const [k] of pairs) keys.push(k);
console.log(keys.join(","));
"#
        ),
        vec!["a,b"]
    );
}

#[test]
fn destructure_object_in_for_of() {
    assert_eq!(
        run_js(
            r#"
const people = [{ name: "X", age: 1 }, { name: "Y", age: 2 }];
const names = [];
for (const { name } of people) names.push(name);
console.log(names.join(","));
"#
        ),
        vec!["X,Y"]
    );
}

#[test]
fn destructure_map_entries() {
    assert_eq!(
        run_js(
            r#"
const m = new Map([["a", 1], ["b", 2]]);
const parts = [];
for (const [k, v] of m) parts.push(k + ":" + v);
console.log(parts.join(","));
"#
        ),
        vec!["a:1,b:2"]
    );
}

#[test]
fn array_destructure_iterator_protocol() {
    assert_eq!(
        run_js(
            r#"
function* gen() { yield 10; yield 20; yield 30; }
const [x, , z] = gen();
console.log(x, z);
"#
        ),
        vec!["10 30"]
    );
}

#[test]
fn object_destructure_prototype_not_included() {
    assert_eq!(
        run_js(
            r#"
const obj = Object.create({ inherited: true });
obj.own = "yes";
const { own, inherited } = obj;
console.log(own);
console.log(inherited);
"#
        ),
        vec!["yes", "true"]
    );
}

#[test]
fn destructure_with_nullish_coalescing_default() {
    assert_eq!(
        run_js(
            r#"
const config = {};
const { timeout = 5000, retries = 3 } = config;
console.log(timeout, retries);
"#
        ),
        vec!["5000 3"]
    );
}

#[test]
fn nested_destructure_function_return() {
    assert_eq!(
        run_js(
            r#"
function getUser() {
  return { id: 1, address: { city: "NYC", zip: "10001" } };
}
const { address: { city } } = getUser();
console.log(city);
"#
        ),
        vec!["NYC"]
    );
}
