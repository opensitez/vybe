use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// ECMAScript: Operators — arithmetic, bitwise, logical,
// assignment, comparison, special operators
// ═══════════════════════════════════════════════════════════

// ── Arithmetic ─────────────────────────────────────────────

#[test]
fn exponentiation() {
    let out = run_js("console.log(2 ** 10);");
    assert_eq!(out, vec!["1024"]);
}

#[test]
fn exponentiation_assign() {
    let out = run_js(r#"
let x = 3;
x **= 3;
console.log(x);
"#);
    assert_eq!(out, vec!["27"]);
}

#[test]
fn modulo() {
    let out = run_js("console.log(17 % 5);");
    assert_eq!(out, vec!["2"]);
}

#[test]
fn unary_plus() {
    let out = run_js(r#"console.log(+"42");"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn unary_negation() {
    let out = run_js("console.log(-42);");
    assert_eq!(out, vec!["-42"]);
}

#[test]
fn prefix_increment() {
    let out = run_js(r#"
let x = 5;
console.log(++x);
console.log(x);
"#);
    assert_eq!(out, vec!["6", "6"]);
}

#[test]
fn postfix_increment() {
    let out = run_js(r#"
let x = 5;
console.log(x++);
console.log(x);
"#);
    assert_eq!(out, vec!["5", "6"]);
}

#[test]
fn prefix_decrement() {
    let out = run_js(r#"
let x = 5;
console.log(--x);
console.log(x);
"#);
    assert_eq!(out, vec!["4", "4"]);
}

#[test]
fn postfix_decrement() {
    let out = run_js(r#"
let x = 5;
console.log(x--);
console.log(x);
"#);
    assert_eq!(out, vec!["5", "4"]);
}

// ── Bitwise ────────────────────────────────────────────────

#[test]
fn bitwise_and() {
    let out = run_js("console.log(0b1100 & 0b1010);");
    assert_eq!(out, vec!["8"]);
}

#[test]
fn bitwise_or() {
    let out = run_js("console.log(0b1100 | 0b1010);");
    assert_eq!(out, vec!["14"]);
}

#[test]
fn bitwise_xor() {
    let out = run_js("console.log(0b1100 ^ 0b1010);");
    assert_eq!(out, vec!["6"]);
}

#[test]
fn bitwise_not() {
    let out = run_js("console.log(~0);");
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn left_shift() {
    let out = run_js("console.log(1 << 4);");
    assert_eq!(out, vec!["16"]);
}

#[test]
fn right_shift() {
    let out = run_js("console.log(16 >> 2);");
    assert_eq!(out, vec!["4"]);
}

#[ignore]
#[test]
fn unsigned_right_shift() {
    let out = run_js("console.log(-1 >>> 0);");
    assert_eq!(out, vec!["4294967295"]);
}

// ── Comparison ─────────────────────────────────────────────

#[test]
fn strict_equality() {
    let out = run_js(r#"
console.log(1 === 1);
console.log(1 === "1");
console.log(null === undefined);
"#);
    assert_eq!(out, vec!["true", "false", "false"]);
}

#[test]
fn strict_inequality() {
    let out = run_js(r#"
console.log(1 !== 2);
console.log(1 !== "1");
"#);
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn loose_equality() {
    let out = run_js(r#"
console.log(1 == 1);
console.log(null == undefined);
"#);
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn comparison_operators() {
    let out = run_js(r#"
console.log(1 < 2);
console.log(2 > 1);
console.log(1 <= 1);
console.log(1 >= 1);
"#);
    assert_eq!(out, vec!["true", "true", "true", "true"]);
}

// ── Logical ────────────────────────────────────────────────

#[test]
fn logical_and() {
    let out = run_js("console.log(true && false);");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn logical_or() {
    let out = run_js("console.log(false || true);");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn logical_not() {
    let out = run_js("console.log(!true);");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn short_circuit_and() {
    let out = run_js(r#"
let x = 0;
false && (x = 1);
console.log(x);
"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn short_circuit_or() {
    let out = run_js(r#"
let x = 0;
true || (x = 1);
console.log(x);
"#);
    assert_eq!(out, vec!["0"]);
}

// ── Nullish & Optional ─────────────────────────────────────

#[test]
fn nullish_coalescing() {
    let out = run_js(r#"
let a = null;
let b = undefined;
let c = 0;
console.log(a ?? "default");
console.log(b ?? "default");
console.log(c ?? "default");
"#);
    assert_eq!(out, vec!["default", "default", "0"]);
}

#[test]
fn optional_chaining_property() {
    let out = run_js(r#"
const obj = { a: { b: 42 } };
console.log(obj?.a?.b);
console.log(obj?.c?.d);
"#);
    assert_eq!(out[0], "42");
}

#[test]
fn optional_chaining_method() {
    let out = run_js(r#"
const obj = {
    greet() { return "hello"; }
};
console.log(obj?.greet());
console.log(obj?.missing?.());
"#);
    assert_eq!(out[0], "hello");
}

// ── Logical Assignment (ES2021) ────────────────────────────

#[test]
fn logical_and_assign() {
    let out = run_js(r#"
let a = 1;
let b = 0;
a &&= 2;
b &&= 2;
console.log(a);
console.log(b);
"#);
    assert_eq!(out, vec!["2", "0"]);
}

#[test]
fn logical_or_assign() {
    let out = run_js(r#"
let a = 0;
let b = 1;
a ||= 42;
b ||= 42;
console.log(a);
console.log(b);
"#);
    assert_eq!(out, vec!["42", "1"]);
}

#[test]
fn nullish_assign() {
    let out = run_js(r#"
let a = null;
let b = 0;
a ??= 42;
b ??= 42;
console.log(a);
console.log(b);
"#);
    assert_eq!(out, vec!["42", "0"]);
}

// ── Compound Assignment ────────────────────────────────────

#[test]
fn compound_assign_all() {
    let out = run_js(r#"
let x = 10;
x += 5; console.log(x);
x -= 3; console.log(x);
x *= 2; console.log(x);
x /= 4; console.log(x);
x %= 5; console.log(x);
"#);
    assert_eq!(out, vec!["15", "12", "24", "6", "1"]);
}

#[test]
fn bitwise_assign() {
    let out = run_js(r#"
let x = 0xFF;
x &= 0x0F;
console.log(x);
x |= 0xF0;
console.log(x);
x ^= 0xFF;
console.log(x);
"#);
    assert_eq!(out, vec!["15", "255", "0"]);
}

#[test]
fn shift_assign() {
    let out = run_js(r#"
let x = 1;
x <<= 4;
console.log(x);
x >>= 2;
console.log(x);
"#);
    assert_eq!(out, vec!["16", "4"]);
}

// ── Ternary ────────────────────────────────────────────────

#[test]
fn ternary_basic() {
    let out = run_js(r#"
const x = 5;
console.log(x > 3 ? "big" : "small");
console.log(x > 10 ? "big" : "small");
"#);
    assert_eq!(out, vec!["big", "small"]);
}

#[test]
fn ternary_nested() {
    let out = run_js(r#"
const x = 5;
const result = x > 10 ? "big" : x > 3 ? "medium" : "small";
console.log(result);
"#);
    assert_eq!(out, vec!["medium"]);
}

// ── typeof ─────────────────────────────────────────────────

#[test]
fn typeof_all_types() {
    let out = run_js(r#"
console.log(typeof 42);
console.log(typeof "hello");
console.log(typeof true);
console.log(typeof undefined);
console.log(typeof null);
console.log(typeof {});
console.log(typeof function(){});
"#);
    assert_eq!(out[0], "number");
    assert_eq!(out[1], "string");
    assert_eq!(out[2], "boolean");
    assert_eq!(out[3], "undefined");
    // typeof null === "object" is a JS quirk
    assert_eq!(out[4], "object");
    assert_eq!(out[5], "object");
    assert_eq!(out[6], "function");
}

// ── instanceof ─────────────────────────────────────────────

#[test]
fn instanceof_basic() {
    let out = run_js(r#"
class Foo {}
const f = new Foo();
console.log(f instanceof Foo);
"#);
    assert_eq!(out, vec!["true"]);
}

// ── in operator ────────────────────────────────────────────

#[test]
fn in_operator() {
    let out = run_js(r#"
const obj = { a: 1, b: 2 };
console.log("a" in obj);
console.log("c" in obj);
"#);
    assert_eq!(out, vec!["true", "false"]);
}

// ── delete ─────────────────────────────────────────────────

#[test]
fn delete_property() {
    let out = run_js(r#"
const obj = { a: 1, b: 2 };
delete obj.a;
console.log("a" in obj);
console.log(obj.b);
"#);
    assert_eq!(out, vec!["false", "2"]);
}

// ── void ───────────────────────────────────────────────────

#[ignore]
#[test]
fn void_operator() {
    let out = run_js(r#"
console.log(void 0);
"#);
    assert_eq!(out, vec!["undefined"]);
}

// ── Comma operator ─────────────────────────────────────────

#[test]
fn comma_operator() {
    let out = run_js(r#"
let x = (1, 2, 3);
console.log(x);
"#);
    assert_eq!(out, vec!["3"]);
}
