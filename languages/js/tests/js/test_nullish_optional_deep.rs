/// Nullish coalescing and optional chaining interaction, edge cases
use super::helpers::run_js;

#[test]
fn nullish_coalescing_null_returns_right() {
    assert_eq!(
        run_js(
            r#"
console.log(null ?? "default");
"#
        ),
        vec!["default"]
    );
}

#[test]
fn nullish_coalescing_undefined_returns_right() {
    assert_eq!(
        run_js(
            r#"
console.log(undefined ?? "fallback");
"#
        ),
        vec!["fallback"]
    );
}

#[test]
fn nullish_coalescing_zero_returns_left() {
    assert_eq!(
        run_js(
            r#"
console.log(0 ?? "ignored");
console.log("" ?? "ignored");
console.log(false ?? "ignored");
"#
        ),
        vec!["0", "", "false"]
    );
}

#[test]
fn nullish_coalescing_short_circuits_rhs() {
    assert_eq!(
        run_js(
            r#"
let called = false;
const side = () => { called = true; return "right"; };
const result = "left" ?? side();
console.log(result);
console.log(called);
"#
        ),
        vec!["left", "false"]
    );
}

#[test]
fn optional_chaining_on_null() {
    assert_eq!(
        run_js(
            r#"
const obj = null;
console.log(obj?.foo);
console.log(obj?.foo?.bar);
"#
        ),
        vec!["undefined", "undefined"]
    );
}

#[test]
fn optional_chaining_on_undefined() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: undefined };
console.log(obj.a?.b);
"#
        ),
        vec!["undefined"]
    );
}

#[test]
fn optional_chaining_with_nullish_default() {
    assert_eq!(
        run_js(
            r#"
const user = null;
const name = user?.name ?? "Guest";
console.log(name);
"#
        ),
        vec!["Guest"]
    );
}

#[test]
fn optional_call_on_undefined_method() {
    assert_eq!(
        run_js(
            r#"
const obj = {};
console.log(obj.method?.());
"#
        ),
        vec!["undefined"]
    );
}

#[test]
fn optional_chain_preserves_zero_and_false() {
    assert_eq!(
        run_js(
            r#"
const obj = { value: 0 };
console.log(obj?.value ?? "default");
const obj2 = { value: false };
console.log(obj2?.value ?? "default");
"#
        ),
        vec!["0", "false"]
    );
}

#[test]
fn optional_chaining_on_array_element() {
    assert_eq!(
        run_js(
            r#"
const arr = [null, { name: "Alice" }];
console.log(arr[0]?.name);
console.log(arr[1]?.name);
"#
        ),
        vec!["undefined", "Alice"]
    );
}

#[test]
fn chained_nullish_coalescing() {
    assert_eq!(
        run_js(
            r#"
const a = null, b = null, c = "found";
console.log(a ?? b ?? c);
"#
        ),
        vec!["found"]
    );
}

#[test]
fn optional_chaining_function_call_with_args() {
    assert_eq!(
        run_js(
            r#"
const obj = {
    greet: (name) => "Hello " + name
};
console.log(obj.greet?.("World"));
console.log(obj.missing?.("World"));
"#
        ),
        vec!["Hello World", "undefined"]
    );
}

#[test]
fn optional_and_nullish_combination_complex() {
    assert_eq!(
        run_js(
            r#"
const config = {
    settings: null,
    timeout: 0,
};
const timeout = config.settings?.timeout ?? config.timeout ?? 3000;
console.log(timeout);
"#
        ),
        vec!["0"]
    );
}
