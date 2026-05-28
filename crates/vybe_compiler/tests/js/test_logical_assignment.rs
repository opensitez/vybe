/// Logical assignment operators — &&=, ||=, ??= (ES2021)

use super::helpers::run_js;

#[test]
fn and_assign_true_left_assigns() {
    assert_eq!(run_js(r#"
let x = 1;
x &&= 42;
console.log(x);
"#), vec!["42"]);
}

#[test]
fn and_assign_false_left_no_assign() {
    assert_eq!(run_js(r#"
let x = 0;
x &&= 42;
console.log(x);
"#), vec!["0"]);
}

#[test]
fn or_assign_false_left_assigns() {
    assert_eq!(run_js(r#"
let x = 0;
x ||= 99;
console.log(x);
"#), vec!["99"]);
}

#[test]
fn or_assign_true_left_no_assign() {
    assert_eq!(run_js(r#"
let x = 5;
x ||= 99;
console.log(x);
"#), vec!["5"]);
}

#[test]
fn nullish_assign_null_assigns() {
    assert_eq!(run_js(r#"
let x = null;
x ??= "default";
console.log(x);
"#), vec!["default"]);
}

#[test]
fn nullish_assign_undefined_assigns() {
    assert_eq!(run_js(r#"
let x;
x ??= "fallback";
console.log(x);
"#), vec!["fallback"]);
}

#[test]
fn nullish_assign_zero_no_assign() {
    assert_eq!(run_js(r#"
let x = 0;
x ??= 99;
console.log(x);
"#), vec!["0"]);
}

#[test]
fn nullish_assign_empty_string_no_assign() {
    assert_eq!(run_js(r#"
let x = "";
x ??= "filled";
console.log(x);
"#), vec![""]);
}

#[test]
fn and_assign_evaluates_rhs_only_when_truthy() {
    assert_eq!(run_js(r#"
let calls = 0;
let x = true;
x &&= (calls++, false);
console.log(calls);
let y = false;
y &&= (calls++, true);
console.log(calls);
"#), vec!["1", "1"]);
}

#[test]
fn or_assign_short_circuits_rhs() {
    assert_eq!(run_js(r#"
let calls = 0;
let x = 1;
x ||= (calls++, 99);
console.log(calls);
let y = 0;
y ||= (calls++, 99);
console.log(calls);
"#), vec!["0", "1"]);
}

#[test]
fn nullish_assign_short_circuits_rhs() {
    assert_eq!(run_js(r#"
let calls = 0;
let x = 0; // not null/undefined
x ??= (calls++, 99);
console.log(calls);
let y = null;
y ??= (calls++, 99);
console.log(calls);
"#), vec!["0", "1"]);
}

#[test]
fn logical_assign_on_object_property() {
    assert_eq!(run_js(r#"
const obj = { a: null, b: 1, c: 0 };
obj.a ??= "filled";
obj.b &&= 42;
obj.c ||= 99;
console.log(obj.a);
console.log(obj.b);
console.log(obj.c);
"#), vec!["filled", "42", "99"]);
}

#[test]
fn chained_logical_assignment() {
    assert_eq!(run_js(r#"
let a = null, b = null, c = "found";
a ??= b ??= c;
console.log(a);
console.log(b);
"#), vec!["found", "found"]);
}
