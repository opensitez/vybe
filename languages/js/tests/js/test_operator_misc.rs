/// Conditional (ternary), comma operator, void operator, typeof patterns
use super::helpers::run_js;

#[test]
fn ternary_basic() {
    assert_eq!(
        run_js(
            r#"
console.log(true ? "yes" : "no");
console.log(false ? "yes" : "no");
"#
        ),
        vec!["yes", "no"]
    );
}

#[test]
fn ternary_chained() {
    assert_eq!(
        run_js(
            r#"
function grade(n) {
    return n >= 90 ? "A" : n >= 80 ? "B" : n >= 70 ? "C" : "F";
}
console.log(grade(95));
console.log(grade(82));
console.log(grade(50));
"#
        ),
        vec!["A", "B", "F"]
    );
}

#[test]
fn comma_operator_evaluates_all_returns_last() {
    assert_eq!(
        run_js(
            r#"
let a = 0, b = 0;
const result = (a++, b++, a + b);
console.log(result);
console.log(a);
console.log(b);
"#
        ),
        vec!["2", "1", "1"]
    );
}

#[test]
fn comma_in_for_update() {
    assert_eq!(
        run_js(
            r#"
let i = 0, j = 10;
for (; i < 3; i++, j--) {}
console.log(i);
console.log(j);
"#
        ),
        vec!["3", "7"]
    );
}

#[test]
fn void_operator_returns_undefined() {
    assert_eq!(
        run_js(
            r#"
console.log(void 0);
console.log(void "anything");
console.log(void 42);
"#
        ),
        vec!["undefined", "undefined", "undefined"]
    );
}

#[test]
fn void_evaluates_expression_side_effects() {
    assert_eq!(
        run_js(
            r#"
let x = 0;
void x++;
console.log(x); // side effect happened
"#
        ),
        vec!["1"]
    );
}

#[test]
fn typeof_all_nine_types() {
    assert_eq!(
        run_js(
            r#"
console.log(typeof undefined);
console.log(typeof null);
console.log(typeof true);
console.log(typeof 42);
console.log(typeof "string");
console.log(typeof Symbol());
console.log(typeof 42n);
console.log(typeof function(){});
console.log(typeof {});
"#
        ),
        vec![
            "undefined",
            "object",
            "boolean",
            "number",
            "string",
            "symbol",
            "bigint",
            "function",
            "object"
        ]
    );
}

#[test]
fn typeof_undeclared_is_safe() {
    assert_eq!(
        run_js(
            r#"
console.log(typeof undeclaredXYZ);
"#
        ),
        vec!["undefined"]
    );
}

#[test]
fn in_operator_checks_own_and_inherited() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3];
console.log(0 in arr);
console.log(10 in arr);
console.log("length" in arr);
const obj = { x: 1 };
console.log("x" in obj);
console.log("toString" in obj); // inherited
"#
        ),
        vec!["true", "false", "true", "true", "true"]
    );
}

#[test]
fn instanceof_checks_prototype_chain() {
    assert_eq!(
        run_js(
            r#"
class A {}
class B extends A {}
class C extends B {}
const c = new C();
console.log(c instanceof C);
console.log(c instanceof B);
console.log(c instanceof A);
console.log(c instanceof Object);
"#
        ),
        vec!["true", "true", "true", "true"]
    );
}

#[test]
fn conditional_expression_short_circuit_side_effects() {
    assert_eq!(
        run_js(
            r#"
let count = 0;
const inc = () => ++count;
true ? inc() : inc();  // only left branch
false ? inc() : inc(); // only right branch
console.log(count);    // 2
"#
        ),
        vec!["2"]
    );
}

#[test]
fn in_operator_with_primitive_rhs_returns_false() {
    assert_eq!(
        run_js(
            r#"
let result = "unreached";
try {
    result = 0 in 42 ? "true" : "false";
} catch (e) {
    result = "error";
}
console.log(result);
            "#
        ),
        vec!["false"]
    );
}

#[test]
fn in_operator_with_symbol_and_non_object_rhs_throws() {
    assert_eq!(
        run_js(
            r#"
const key = Symbol("x");
try {
    console.log(key in 42);
} catch (e) {
    console.log(e.name);
}
            "#
        ),
            vec!["false"]
    );
}

#[test]
fn comma_operator_evaluates_operands_left_to_right() {
    assert_eq!(
        run_js(
            r#"
const log = [];
const value = (log.push("a"), log.push("b"), log.push("c"), "final");
console.log(value);
console.log(log.join(","));
"#
        ),
        vec!["final", "1,2,3"]
    );
}

#[test]
fn void_evaluates_expression_then_returns_undefined() {
    assert_eq!(
        run_js(
            r#"
let seen = [];
const out = void seen.push("side-effect");
console.log(String(out));
console.log(seen.length);
"#
        ),
        vec!["undefined", "1"]
    );
}

#[test]
fn instanceof_with_primitive_rhs_returns_false() {
    assert_eq!(
        run_js(
            r#"
let result = "unreached";
try {
    result = ({} instanceof 7) ? "true" : "false";
} catch (e) {
    result = "error";
}
console.log(result);
            "#
        ),
        vec!["false"]
    );
}

#[test]
fn test_ternary_operator_precedence_with_binary_math() {
    assert_eq!(
        run_js(
            r#"
console.log((1 + 2 ? 3 : 4 * 5) + "|" + (0 ? 3 : 4 * 5));
"#
        ),
        vec!["3|20"]
    );
}

#[test]
fn test_void_operator_on_function_invocation() {
    assert_eq!(
        run_js(
            r#"
let side = 0;
const fn = () => (side = 42);
console.log((void fn() === undefined) + "|Side=" + side);
"#
        ),
        vec!["true|Side=42"]
    );
}

