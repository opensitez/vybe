//! Ruby multi-value tuple returns. Ruby's `return a, b` + `x, y = f()`
//! idiom maps to WASM multi-value via the same pre-scan used for Python;
//! non-destructure callers (`r = f()`) get the N values auto-repacked
//! into an array so they behave like Ruby's native array-return.

use super::helpers::run_ruby;

#[test]
fn two_value_return_and_destructure() {
    let out = run_ruby(
        r##"
def swap(a, b)
    return b, a
end

x, y = swap(1, 2)
puts x
puts y
"##,
    );
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn three_value_return_and_destructure() {
    let out = run_ruby(
        r##"
def rgb
    return 10, 20, 30
end

r, g, b = rgb()
puts r
puts g
puts b
"##,
    );
    assert_eq!(out, vec!["10", "20", "30"]);
}

#[test]
fn multi_value_function_used_as_array_repacks() {
    // Non-destructure callers keep Ruby's native semantics — the N
    // multi-value returns are re-packed into an array so `r[0]`/`r[1]`
    // work as expected.
    let out = run_ruby(
        r#"
def pair
    return 10, 20
end

r = pair()
puts r[0]
puts r[1]
"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}
