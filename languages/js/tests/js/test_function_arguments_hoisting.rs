/// Function hoisting in blocks, arguments object sloppy mode, callee patterns
use super::helpers::run_js;

#[test]
fn arguments_length_reflects_call() {
    assert_eq!(
        run_js(
            r#"
function f() { return arguments.length; }
console.log(f(1, 2, 3));
console.log(f());
"#
        ),
        vec!["3", "0"]
    );
}

#[test]
fn arguments_index_access() {
    assert_eq!(
        run_js(
            r#"
function sum() {
    let s = 0;
    for (let i = 0; i < arguments.length; i++) s += arguments[i];
    return s;
}
console.log(sum(1, 2, 3, 4));
"#
        ),
        vec!["10"]
    );
}

#[test]
fn arguments_is_array_like_not_array() {
    assert_eq!(
        run_js(
            r#"
function f() { return Array.isArray(arguments); }
console.log(f(1, 2));
"#
        ),
        vec!["false"]
    );
}

#[test]
fn arguments_to_array_via_spread() {
    assert_eq!(
        run_js(
            r#"
function f() { return [...arguments].join(","); }
console.log(f("a", "b", "c"));
"#
        ),
        vec!["a,b,c"]
    );
}

#[test]
fn arguments_not_in_arrow_function() {
    assert_eq!(
        run_js(
            r#"
function outer() {
    const inner = () => arguments[0]; // arrow captures outer's arguments
    return inner();
}
console.log(outer(42));
"#
        ),
        vec!["42"]
    );
}

#[test]
fn function_name_property() {
    assert_eq!(
        run_js(
            r#"
function named() {}
const anon = function() {};
const arrow = () => {};
console.log(named.name);
console.log(anon.name);
console.log(arrow.name);
"#
        ),
        vec!["named", "anon", "arrow"]
    );
}

#[test]
fn function_length_excludes_defaults_and_rest() {
    assert_eq!(
        run_js(
            r#"
function f(a, b, c = 1, ...rest) {}
console.log(f.length); // only a, b count
"#
        ),
        vec!["2"]
    );
}

#[test]
fn function_in_block_behavior() {
    assert_eq!(
        run_js(
            r#"
// Block-scoped function declaration (non-strict behavior: hoisted as var)
console.log(typeof blockFn);
{
    function blockFn() { return "inside"; }
}
console.log(typeof blockFn);
"#
        ),
        vec!["undefined", "function"]
    );
}

#[test]
fn immediately_invoked_arrow() {
    assert_eq!(
        run_js(
            r#"
const result = ((x, y) => x + y)(3, 4);
console.log(result);
"#
        ),
        vec!["7"]
    );
}

#[test]
fn function_expression_name_in_own_scope() {
    assert_eq!(
        run_js(
            r#"
const factorial = function fact(n) {
    return n <= 1 ? 1 : n * fact(n - 1);
};
console.log(factorial(5));
console.log(typeof fact); // not accessible outside
"#
        ),
        vec!["120", "undefined"]
    );
}

#[test]
fn rest_parameter_is_real_array() {
    assert_eq!(
        run_js(
            r#"
function f(...args) { return Array.isArray(args); }
console.log(f(1, 2, 3));
"#
        ),
        vec!["true"]
    );
}

#[test]
fn default_param_uses_previous_param() {
    assert_eq!(
        run_js(
            r#"
function greet(name, msg = "Hello " + name) {
    return msg;
}
console.log(greet("World"));
console.log(greet("World", "Hi World"));
"#
        ),
        vec!["Hello World", "Hi World"]
    );
}

#[test]
fn default_param_not_evaluated_if_provided() {
    assert_eq!(
        run_js(
            r#"
let count = 0;
function f(x = (count++, 0)) { return x; }
f(5);
console.log(count);
f();
console.log(count);
"#
        ),
        vec!["0", "1"]
    );
}

#[test]
fn default_param_evaluated_when_passed_undefined() {
    assert_eq!(
        run_js(
            r#"
function f(x = "default") { return x; }
console.log(f(undefined));
console.log(f(null));
"#
        ),
        vec!["default", "null"]
    );
}

