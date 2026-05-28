/// Array destructuring advanced — defaults, skip, nested, iterator protocol, rest

use super::helpers::run_js;

#[test]
fn destructure_skip_elements() {
    assert_eq!(run_js(r#"
const [a, , b, , c] = [1, 2, 3, 4, 5];
console.log(a);
console.log(b);
console.log(c);
"#), vec!["1", "3", "5"]);
}

#[test]
fn destructure_default_values() {
    assert_eq!(run_js(r#"
const [a = 10, b = 20, c = 30] = [1, 2];
console.log(a);
console.log(b);
console.log(c);
"#), vec!["1", "2", "30"]);
}

#[test]
fn destructure_with_rest() {
    assert_eq!(run_js(r#"
const [head, ...tail] = [1, 2, 3, 4, 5];
console.log(head);
console.log(tail.join(","));
"#), vec!["1", "2,3,4,5"]);
}

#[test]
fn destructure_nested_arrays() {
    assert_eq!(run_js(r#"
const [[a, b], [c, d]] = [[1, 2], [3, 4]];
console.log(a);
console.log(b);
console.log(c);
console.log(d);
"#), vec!["1", "2", "3", "4"]);
}

#[test]
fn destructure_swap_variables() {
    assert_eq!(run_js(r#"
let x = 1, y = 2;
[x, y] = [y, x];
console.log(x);
console.log(y);
"#), vec!["2", "1"]);
}

#[test]
fn destructure_from_function_return() {
    assert_eq!(run_js(r#"
function minMax(arr) {
    return [Math.min(...arr), Math.max(...arr)];
}
const [min, max] = minMax([3, 1, 4, 1, 5, 9]);
console.log(min);
console.log(max);
"#), vec!["1", "9"]);
}

#[test]
fn destructure_from_string() {
    assert_eq!(run_js(r#"
const [a, b, c] = "hello";
console.log(a);
console.log(b);
console.log(c);
"#), vec!["h", "e", "l"]);
}

#[test]
fn destructure_generator() {
    assert_eq!(run_js(r#"
function* gen() { yield 1; yield 2; yield 3; }
const [a, b] = gen();
console.log(a);
console.log(b);
"#), vec!["1", "2"]);
}

#[test]
fn destructure_map_entries_pattern() {
    assert_eq!(run_js(r#"
const map = new Map([["x", 1], ["y", 2]]);
for (const [key, val] of map) {
    console.log(key + "=" + val);
}
"#), vec!["x=1", "y=2"]);
}

#[test]
fn destructure_object_rest() {
    assert_eq!(run_js(r#"
const { a, b, ...rest } = { a: 1, b: 2, c: 3, d: 4 };
console.log(a);
console.log(b);
console.log(Object.keys(rest).sort().join(","));
"#), vec!["1", "2", "c,d"]);
}

#[test]
fn destructure_rename_and_default() {
    assert_eq!(run_js(r#"
const { x: myX = 10, y: myY = 20 } = { x: 5 };
console.log(myX);
console.log(myY);
"#), vec!["5", "20"]);
}

#[test]
fn nested_object_array_mixed() {
    assert_eq!(run_js(r#"
const { name, scores: [first, ...rest] } = { name: "Alice", scores: [95, 88, 72] };
console.log(name);
console.log(first);
console.log(rest.join(","));
"#), vec!["Alice", "95", "88,72"]);
}
