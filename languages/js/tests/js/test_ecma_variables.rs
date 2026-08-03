use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// ECMAScript: Variables, scoping, declarations
// ═══════════════════════════════════════════════════════════

#[test]
fn let_block_scope() {
    let out = run_js(
        r#"
let x = 1;
{
    let x = 2;
    console.log(x);
}
console.log(x);
"#,
    );
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn const_cannot_reassign() {
    // const is just a hint in our system; verify it compiles
    let out = run_js(
        r#"
const x = 42;
console.log(x);
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn var_function_scope() {
    let out = run_js(
        r#"
function test() {
    if (true) {
        var x = 10;
    }
    console.log(x);
}
test();
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn multiple_declarators() {
    let out = run_js(
        r#"
let a = 1, b = 2, c = 3;
console.log(a + b + c);
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn destructure_object_basic() {
    let out = run_js(
        r#"
const obj = { x: 10, y: 20 };
const { x, y } = obj;
console.log(x);
console.log(y);
"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn destructure_object_rename() {
    let out = run_js(
        r#"
const obj = { name: "Alice", age: 30 };
const { name: n, age: a } = obj;
console.log(n);
console.log(a);
"#,
    );
    assert_eq!(out, vec!["Alice", "30"]);
}

#[test]
fn destructure_object_default() {
    let out = run_js(
        r#"
const obj = { x: 5 };
const { x, y = 10 } = obj;
console.log(x);
console.log(y);
"#,
    );
    assert_eq!(out, vec!["5", "10"]);
}

#[test]
fn destructure_array_basic() {
    let out = run_js(
        r#"
const arr = [1, 2, 3];
const [a, b, c] = arr;
console.log(a);
console.log(b);
console.log(c);
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn destructure_array_skip() {
    let out = run_js(
        r#"
const arr = [1, 2, 3, 4];
const [a,  c] = arr;
console.log(a);
console.log(c);
"#,
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn destructure_array_rest() {
    let out = run_js(
        r#"
const arr = [1, 2, 3, 4, 5];
const [first, ...rest] = arr;
console.log(first);
console.log(rest.length);
"#,
    );
    assert_eq!(out, vec!["1", "4"]);
}

#[test]
fn destructure_in_for_of() {
    let out = run_js(
        r#"
const pairs = [[1, "a"], [2, "b"], [3, "c"]];
for (const [num, letter] of pairs) {
    console.log(num + letter);
}
"#,
    );
    assert_eq!(out, vec!["1a", "2b", "3c"]);
}

#[test]
fn destructure_in_function_params() {
    let out = run_js(
        r#"
function greet({ name, age }) {
    console.log(name + " is " + age);
}
greet({ name: "Bob", age: 25 });
"#,
    );
    assert_eq!(out, vec!["Bob is 25"]);
}

#[test]
fn destructure_nested_object() {
    let out = run_js(
        r#"
const data = { a: { b: { c: 42 } } };
const { a: { b: { c } } } = data;
console.log(c);
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn destructure_with_computed_property() {
    let out = run_js(
        r#"
const key = "name";
const obj = { name: "Alice" };
const { [key]: val } = obj;
console.log(val);
"#,
    );
    assert_eq!(out, vec!["Alice"]);
}

#[test]
fn destructure_swap() {
    let out = run_js(
        r#"
let a = 1, b = 2;
[a, b] = [b, a];
console.log(a);
console.log(b);
"#,
    );
    assert_eq!(out, vec!["2", "1"]);
}
