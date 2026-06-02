//! Python multi-value tuple returns — verifies the compiler routes
//! uniform-arity `return a, b` through the WASM multi-value ABI and the
//! VM RETURN pops N values back onto the caller's stack so
//! `a, b = f()` destructuring reads them off directly.

use super::helpers::{run_python, run_python_one};

#[test]
fn two_value_return_and_destructure() {
    let out = run_python_one(
        r#"
def swap(a, b):
    return b, a

x, y = swap(1, 2)
print(x, y)
"#,
    );
    assert_eq!(out, "2 1");
}

#[test]
fn three_value_return_and_destructure() {
    let out = run_python_one(
        r#"
def rgb():
    return 10, 20, 30

r, g, b = rgb()
print(r, g, b)
"#,
    );
    assert_eq!(out, "10 20 30");
}

#[test]
fn multi_value_through_branches() {
    // Function with two tuple-return paths of matching arity → multi-value.
    let out = run_python(
        r#"
def pick(n):
    if n > 0:
        return n, n * 2
    return 0, 0

a, b = pick(5)
print(a, b)
c, d = pick(0)
print(c, d)
"#,
    );
    assert_eq!(out, vec!["5 10", "0 0"]);
}

#[test]
fn multi_value_function_used_as_single_tuple_repacks() {
    // A multi-return function is sometimes used without destructuring
    // (`r = f()`, then `r[0]` / `r[1]`). The caller must re-pack the
    // N multi-value results into a single tuple so those use sites
    // see the expected Python semantics.
    let out = run_python(
        r#"
def pair():
    return 10, 20

r = pair()
print(r[0])
print(r[1])
"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn non_uniform_returns_stay_single_value() {
    // Mixed tuple + scalar returns must NOT opt into multi-value —
    // the function still returns a single value, and a caller that
    // treats the result as a scalar sees the last pushed value.
    let out = run_python_one(
        r#"
def mixed(flag):
    if flag:
        return (1, 2)
    return 7

v = mixed(False)
print(v)
"#,
    );
    assert_eq!(out, "7");
}
