/// Operators — bitwise ops, exponentiation, in/instanceof advanced,
/// comma operator in loops, ternary chains, typeof checks, void,
/// logical operators with non-boolean values.
use super::helpers::run_js;

// ── bitwise operators ─────────────────────────────────────────────────────────

#[test]
fn bitwise_and_or_xor() {
    assert_eq!(
        run_js(
            r#"
console.log(0b1010 & 0b1100);  // AND: 0b1000 = 8
console.log(0b1010 | 0b1100);  // OR:  0b1110 = 14
console.log(0b1010 ^ 0b1100);  // XOR: 0b0110 = 6
"#
        ),
        vec!["8", "14", "6"]
    );
}

#[test]
fn bitwise_not() {
    assert_eq!(
        run_js(
            r#"
console.log(~0);   // -1
console.log(~1);   // -2
console.log(~-1);  // 0
"#
        ),
        vec!["-1", "-2", "0"]
    );
}

#[test]
fn left_shift_right_shift() {
    assert_eq!(
        run_js(
            r#"
console.log(1 << 4);   // 16
console.log(16 >> 2);  // 4
console.log(-1 >> 1);  // -1 (sign preserved)
console.log(-1 >>> 1); // 2147483647 (unsigned)
"#
        ),
        vec!["16", "4", "-1", "2147483647"]
    );
}

#[test]
fn unsigned_right_shift() {
    assert_eq!(
        run_js(
            r#"
console.log(-1 >>> 0);  // 4294967295 (convert to uint32)
console.log(0xFFFFFFFF >>> 0); // 4294967295
"#
        ),
        vec!["4294967295", "4294967295"]
    );
}

// ── bitwise use cases ─────────────────────────────────────────────────────────

#[test]
fn bitwise_flag_manipulation() {
    assert_eq!(
        run_js(
            r#"
const READ   = 0b001;
const WRITE  = 0b010;
const EXEC   = 0b100;

let perms = READ | WRITE; // 0b011
console.log((perms & READ) !== 0);  // has READ
console.log((perms & EXEC) !== 0);  // has EXEC (no)
perms |= EXEC;                       // add EXEC
console.log((perms & EXEC) !== 0);  // has EXEC (yes)
perms &= ~WRITE;                     // remove WRITE
console.log((perms & WRITE) !== 0); // has WRITE (no)
"#
        ),
        vec!["true", "false", "true", "false"]
    );
}

#[test]
fn double_tilde_converts_to_int32() {
    assert_eq!(
        run_js(
            r#"
console.log(~~3.9);    // 3
console.log(~~-3.9);   // -3
console.log(~~"42");   // 42
"#
        ),
        vec!["3", "-3", "42"]
    );
}

// ── exponentiation ────────────────────────────────────────────────────────────

#[test]
fn exponentiation_operator() {
    assert_eq!(
        run_js(
            r#"
console.log(2 ** 10);
console.log(3 ** 3);
console.log((-2) ** 3);
"#
        ),
        vec!["1024", "27", "-8"]
    );
}

#[test]
fn exponentiation_assignment() {
    assert_eq!(
        run_js(
            r#"
let x = 2;
x **= 8;
console.log(x);
"#
        ),
        vec!["256"]
    );
}

// ── in operator ───────────────────────────────────────────────────────────────

#[test]
fn in_with_array_index() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3];
console.log(0 in arr);
console.log(2 in arr);
console.log(3 in arr);  // out of bounds
"#
        ),
        vec!["true", "true", "false"]
    );
}

#[test]
fn in_with_string_key() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: undefined }; // key exists but value undefined
console.log("a" in obj);
console.log("b" in obj);
console.log(obj.a === undefined); // exists but undefined
"#
        ),
        vec!["true", "false", "true"]
    );
}

// ── instanceof advanced ───────────────────────────────────────────────────────

#[test]
fn instanceof_custom_symbol_hasinstance() {
    assert_eq!(
        run_js(
            r#"
class EvenChecker {
    static [Symbol.hasInstance](n) {
        return typeof n === "number" && n % 2 === 0;
    }
}
console.log(2 instanceof EvenChecker);
console.log(3 instanceof EvenChecker);
console.log(4 instanceof EvenChecker);
"#
        ),
        vec!["true", "false", "true"]
    );
}

// ── logical operators with non-booleans ───────────────────────────────────────

#[test]
fn logical_and_returns_first_falsy_or_last() {
    assert_eq!(
        run_js(
            r#"
console.log(1 && 2 && 3);   // last value
console.log(1 && 0 && 3);   // first falsy
console.log("" && "abc");   // first falsy
"#
        ),
        vec!["3", "0", ""]
    );
}

#[test]
fn logical_or_returns_first_truthy_or_last() {
    assert_eq!(
        run_js(
            r#"
console.log(0 || false || 3);  // first truthy
console.log(0 || false || ""); // all falsy — last value
console.log(1 || 2 || 3);      // first truthy
"#
        ),
        vec!["3", "", "1"]
    );
}

#[test]
fn nullish_coalescing_only_null_undefined() {
    assert_eq!(
        run_js(
            r#"
console.log(0 ?? "default");      // 0 — not null/undefined
console.log("" ?? "default");     // "" — not null/undefined
console.log(false ?? "default");  // false — not null/undefined
console.log(null ?? "default");   // "default"
console.log(undefined ?? "default"); // "default"
"#
        ),
        vec!["0", "", "false", "default", "default"]
    );
}

// ── comma operator ────────────────────────────────────────────────────────────

#[test]
fn comma_in_for_update() {
    assert_eq!(
        run_js(
            r#"
const result = [];
for (let i = 0, j = 10; i < 3; i++, j -= 3) {
    result.push(i + ":" + j);
}
console.log(result.join(","));
"#
        ),
        vec!["0:10,1:7,2:4"]
    );
}

// ── void operator ─────────────────────────────────────────────────────────────

#[test]
fn void_evaluates_then_returns_undefined() {
    assert_eq!(
        run_js(
            r#"
let x = 0;
const result = void (x = 42);
console.log(result);
console.log(x); // side effect happened
"#
        ),
        vec!["undefined", "42"]
    );
}

// ── typeof ────────────────────────────────────────────────────────────────────

#[test]
fn typeof_all_types() {
    assert_eq!(
        run_js(
            r#"
console.log(typeof 42);
console.log(typeof "str");
console.log(typeof true);
console.log(typeof undefined);
console.log(typeof null);
console.log(typeof {});
console.log(typeof []);
console.log(typeof function(){});
console.log(typeof Symbol());
"#
        ),
        vec![
            "number",
            "string",
            "boolean",
            "undefined",
            "object",
            "object",
            "object",
            "function",
            "symbol"
        ]
    );
}

// ── ternary chaining ──────────────────────────────────────────────────────────

#[test]
fn ternary_chain_grade() {
    assert_eq!(
        run_js(
            r#"
function grade(n) {
    return n >= 90 ? "A"
         : n >= 80 ? "B"
         : n >= 70 ? "C"
         : n >= 60 ? "D" : "F";
}
console.log(grade(95));
console.log(grade(82));
console.log(grade(55));
"#
        ),
        vec!["A", "B", "F"]
    );
}

#[test]
fn exponentiation_precedence_and_syntax_boundary() {
    assert_eq!(
        run_js(
            r#"
console.log((2 ** 3) ** 2); // explicit grouping
console.log(2 ** 3 ** 2);    // right-associative exponentiation
console.log("end");
"#
        ),
        vec!["64", "64", "end"]
    );
}

#[test]
fn typeof_undeclared_is_safe_and_tdz_throws() {
    assert_eq!(
        run_js(
            r#"
console.log(typeof doesNotExist);
try {
    eval("{ console.log(typeof tdzVar); let tdzVar = 123; }");
} catch (e) {
    console.log("TDZ");
}
"#
        ),
        vec!["undefined", "TDZ"]
    );
}

#[test]
fn in_operator_considers_prototype_chain() {
    assert_eq!(
        run_js(
            r#"
const proto = { inherited: 1 };
const child = Object.create(proto);
console.log("inherited" in child);
console.log(child.hasOwnProperty("inherited"));
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn instanceof_checks_inheritance_chain() {
    assert_eq!(
        run_js(
            r#"
class Base {}
class Derived extends Base {}
const value = new Derived();
console.log(value instanceof Derived);
console.log(value instanceof Base);
console.log(42 instanceof Base);
"#
        ),
        vec!["true", "true", "false"]
    );
}

#[test]
fn compound_assignment_and_operator_precedence() {
    assert_eq!(
        run_js(
            r#"
let x = 20;
x += 10; // 30
x -= 5;  // 25
x *= 2;  // 50
x /= 5;  // 10
x %= 7;  // 3
x **= 2; // 9
console.log(x);
"#
        ),
        vec!["9"]
    );
}

#[test]
fn precedence_addition_before_shift() {
    assert_eq!(
        run_js(
            r#"
console.log(1 + 2 << 2); // (1 + 2) << 2 = 12
console.log(1 + (2 << 2)); // 9
"#
        ),
        vec!["12", "9"]
    );
}

#[test]
fn chained_comparisons_and_associativity() {
    assert_eq!(
        run_js(
            r#"
console.log(1 + 2 * 3 - 4 / 2);
console.log(1 < 2 < 3); // left-to-right: (1 < 2) -> true -> 1 -> 1 < 3
console.log(3 < 2 < 1); // left-to-right: (3 < 2) -> false -> 0 -> 0 < 1
"#
        ),
        vec!["5", "true", "true"]
    );
}

#[test]
fn division_and_remainder_edge_cases() {
    assert_eq!(
        run_js(
            r#"
console.log(5 / 0);
console.log(-5 / 0);
console.log(0 / 0);
console.log(5 % 0);
"#
        ),
        vec!["Infinity", "-Infinity", "NaN", "NaN"]
    );
}

#[test]
fn division_and_remainder_negative_values() {
    assert_eq!(
        run_js(
            r#"
console.log(5 / -2);
console.log(-5 / 2);
console.log(-5 % 2);
console.log(5 % -2);
console.log((-5 % -2));
"#
        ),
        vec!["-2.5", "-2.5", "-1", "1", "-1"]
    );
}

#[test]
fn remainder_with_mixed_sign_and_float_values() {
    assert_eq!(
        run_js(
            r#"
console.log(10 % 4);
console.log(10 % -4);
console.log(-10 % 4);
console.log(10.5 % 2);
"#
        ),
        vec!["2", "2", "-2", "0.5"]
    );
}

#[test]
fn logical_assignment_short_circuits_rhs() {
    assert_eq!(
        run_js(
            r#"
let value = 5;
let calls = 0;
value ||= (() => {
    calls++;
    return 9;
})();
let ready = 0;
ready ||= (() => {
    calls++;
    return 11;
})();
console.log(value);
console.log(ready);
console.log(calls);
"#
        ),
        vec!["5", "11", "1"]
    );
}

#[test]
fn nullish_assignment_sets_only_when_nullish() {
    assert_eq!(
        run_js(
            r#"
let a;
let b = 0;
a ??= 5;
b ??= 7;
console.log(a);
console.log(b);
"#
        ),
        vec!["5", "0"]
    );
}

#[test]
fn in_operator_requires_rhs_object() {
    assert_eq!(
        run_js(
            r#"
console.log("x" in { x: 1, y: 2 });
console.log("x" in 1);
console.log("x" in Object(1));
"#
        ),
        vec!["true", "false", "false"]
    );
}

#[test]
fn in_operator_rhs_null_or_undefined_does_not_throw_typeerror_in_vm() {
    assert_eq!(
        run_js(
            r#"
let nullCheck = false;
let undefinedCheck = false;
try {
    "x" in null;
} catch (e) {
    nullCheck = e instanceof TypeError;
}
try {
    "x" in undefined;
} catch (e) {
    undefinedCheck = e instanceof TypeError;
}
console.log(`${nullCheck}:${undefinedCheck}`);
            "#
        ),
        vec!["false:false"]
    );
}

#[test]
fn instanceof_rhs_coercion_with_non_function_is_false() {
    assert_eq!(
        run_js(
            r#"
console.log(3 instanceof Number);
console.log(3 instanceof 123);
console.log({} instanceof Number);
console.log({} instanceof Object);
"#
        ),
        vec!["false", "false", "false", "false"]
    );
}

#[test]
fn delete_non_configurable_property_returns_false() {
    assert_eq!(
        run_js(
            r#"
const obj = Object.defineProperty({}, "x", { value: 1, configurable: false });
console.log(delete obj.x);
console.log(obj.x);
"#
        ),
        vec!["false", "1"]
    );
}

#[test]
fn delete_non_existent_property_returns_true() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1 };
console.log(delete obj.missing);
console.log(obj.a);
"#
        ),
        vec!["true", "1"]
    );
}

#[test]
fn instanceof_symbol_hasinstance_truthiness() {
    assert_eq!(
        run_js(
            r#"
class FalseInstanceof {
    static [Symbol.hasInstance]() { return ""; }
}
class TruthyInstanceof {
    static [Symbol.hasInstance]() { return 42; }
}
console.log({} instanceof FalseInstanceof);
console.log({} instanceof TruthyInstanceof);
"#
        ),
        vec!["false", "true"]
    );
}

#[test]
fn in_operator_with_symbol_property_key() {
    assert_eq!(
        run_js(
            r#"
const token = Symbol("token");
const obj = { [token]: 123 };
console.log(token in obj);
console.log(Object.hasOwn(obj, token));
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn bigint_binary_operators_accept_same_type_only() {
    assert_eq!(
        run_js(
            r#"
console.log((1n + 2n).toString());
console.log((1n << 2n).toString());
let addMixed = false;
let shiftMixed = false;
try {
    const _ = 1n + 2;
} catch (e) {
    addMixed = true;
}
try {
    const _ = 1n << 2;
} catch (e) {
    shiftMixed = true;
}
console.log(`${addMixed}:${shiftMixed}`);
"#
        ),
        vec!["3", "4", "true:true"]
    );
}

#[test]
fn nullish_coalescing_on_boolean_and_zero() {
    assert_eq!(
        run_js(
            r#"
console.log((false ?? "fallback"));
console.log((0 ?? "fallback"));
console.log((null ?? "fallback"));
"#
        ),
        vec!["false", "0", "fallback"]
    );
}

#[test]
fn bitwise_shift_counts_are_masked_to_32_bits() {
    assert_eq!(
        run_js(
            r#"
console.log(1 << 32);
console.log(1 << 33);
console.log(1 << 40);
console.log(-1 << 1);
"#
        ),
        vec!["1", "2", "256", "-2"]
    );
}

#[test]
fn bitwise_shift_counts_can_be_negative() {
    assert_eq!(
        run_js(
            r#"
console.log(1 << -1);   // -1 -> 31
console.log(-1 << -1);  // -1 -> 31
console.log(1 >> -1);   // same as >>31
console.log(-1 >> -1);
"#
        ),
        vec!["-2147483648", "-2147483648", "0", "-1"]
    );
}

#[test]
fn arithmetic_addition_string_vs_numeric_coercion() {
    let src = r#"
console.log("1" + 2 + 3);
console.log(1 + "2" + 3);
console.log("1" + (2 + 3));
console.log(1 + 2 + "3");
console.log("1" - 2);
console.log("1" * "2");
"#;
    assert_eq!(run_js(src), vec!["123", "123", "15", "33", "-1", "2"]);
}

#[test]
fn arithmetic_unary_conversion_operations() {
    let src = r#"
console.log(+true);
console.log(+false);
console.log(-"42");
console.log(+null);
console.log(-false);
"#;
    assert_eq!(run_js(src), vec!["1", "0", "-42", "0", "0"]);
}

#[test]
fn test_bigint_relational_comparison_with_infinity() {
    let src = r#"
console.log(`${100n < Infinity}:${-100n > -Infinity}:${100n > -Infinity}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:true"]);
}
