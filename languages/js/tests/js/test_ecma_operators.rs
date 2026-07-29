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
    let out = run_js(
        r#"
let x = 3;
x **= 3;
console.log(x);
"#,
    );
    assert_eq!(out, vec!["27"]);
}

#[test]
fn modulo() {
    let out = run_js("console.log(17 % 5);");
    assert_eq!(out, vec!["2"]);
}

#[test]
fn arithmetic_nan_and_infinity() {
    let out = run_js(
        r#"
console.log(1 / 0);
console.log(-1 / 0);
console.log(0 / 0);
console.log(10 + NaN);
console.log(Infinity - Infinity);
"#,
    );
    assert_eq!(out, vec!["Infinity", "-Infinity", "NaN", "NaN", "NaN"]);
}

#[test]
fn unary_plus() {
    let out = run_js(r#"console.log(+"42");"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn subtraction_with_boolean_and_string_operands() {
    let out = run_js(
        r#"
console.log("10" - true);
console.log(10 - false);
console.log("3" - "1");
console.log(true - false);
"#,
    );
    assert_eq!(out, vec!["9", "10", "2", "1"]);
}

#[test]
fn unary_negation() {
    let out = run_js("console.log(-42);");
    assert_eq!(out, vec!["-42"]);
}

#[test]
fn prefix_increment() {
    let out = run_js(
        r#"
let x = 5;
console.log(++x);
console.log(x);
"#,
    );
    assert_eq!(out, vec!["6", "6"]);
}

#[test]
fn postfix_increment() {
    let out = run_js(
        r#"
let x = 5;
console.log(x++);
console.log(x);
"#,
    );
    assert_eq!(out, vec!["5", "6"]);
}

#[test]
fn prefix_decrement() {
    let out = run_js(
        r#"
let x = 5;
console.log(--x);
console.log(x);
"#,
    );
    assert_eq!(out, vec!["4", "4"]);
}

#[test]
fn postfix_decrement() {
    let out = run_js(
        r#"
let x = 5;
console.log(x--);
console.log(x);
"#,
    );
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

#[test]
fn unsigned_right_shift() {
    let out = run_js("console.log(-1 >>> 0);");
    assert_eq!(out, vec!["4294967295"]);
}

// ── Comparison ─────────────────────────────────────────────

#[test]
fn strict_equality() {
    let out = run_js(
        r#"
console.log(1 === 1);
console.log(1 === "1");
console.log(null === undefined);
"#,
    );
    assert_eq!(out, vec!["true", "false", "false"]);
}

#[test]
fn strict_inequality() {
    let out = run_js(
        r#"
console.log(1 !== 2);
console.log(1 !== "1");
"#,
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn loose_equality() {
    let out = run_js(
        r#"
console.log(1 == 1);
console.log(null == undefined);
"#,
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn comparison_operators() {
    let out = run_js(
        r#"
console.log(1 < 2);
console.log(2 > 1);
console.log(1 <= 1);
console.log(1 >= 1);
"#,
    );
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
    let out = run_js(
        r#"
let x = 0;
false && (x = 1);
console.log(x);
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn short_circuit_or() {
    let out = run_js(
        r#"
let x = 0;
true || (x = 1);
console.log(x);
"#,
    );
    assert_eq!(out, vec!["0"]);
}

// ── Nullish & Optional ─────────────────────────────────────

#[test]
fn nullish_coalescing() {
    let out = run_js(
        r#"
let a = null;
let b = undefined;
let c = 0;
console.log(a ?? "default");
console.log(b ?? "default");
console.log(c ?? "default");
"#,
    );
    assert_eq!(out, vec!["default", "default", "0"]);
}

#[test]
fn optional_chaining_property() {
    let out = run_js(
        r#"
const obj = { a: { b: 42 } };
console.log(obj?.a?.b);
console.log(obj?.c?.d);
"#,
    );
    assert_eq!(out[0], "42");
}

#[test]
fn optional_chaining_method() {
    let out = run_js(
        r#"
const obj = {
    greet() { return "hello"; }
};
console.log(obj?.greet());
console.log(obj?.missing?.());
"#,
    );
    assert_eq!(out[0], "hello");
}

// ── Logical Assignment (ES2021) ────────────────────────────

#[test]
fn logical_and_assign() {
    let out = run_js(
        r#"
let a = 1;
let b = 0;
a &&= 2;
b &&= 2;
console.log(a);
console.log(b);
"#,
    );
    assert_eq!(out, vec!["2", "0"]);
}

#[test]
fn logical_or_assign() {
    let out = run_js(
        r#"
let a = 0;
let b = 1;
a ||= 42;
b ||= 42;
console.log(a);
console.log(b);
"#,
    );
    assert_eq!(out, vec!["42", "1"]);
}

#[test]
fn nullish_assign() {
    let out = run_js(
        r#"
let a = null;
let b = 0;
a ??= 42;
b ??= 42;
console.log(a);
console.log(b);
"#,
    );
    assert_eq!(out, vec!["42", "0"]);
}

// ── Compound Assignment ────────────────────────────────────

#[test]
fn compound_assign_all() {
    let out = run_js(
        r#"
let x = 10;
x += 5; console.log(x);
x -= 3; console.log(x);
x *= 2; console.log(x);
x /= 4; console.log(x);
x %= 5; console.log(x);
"#,
    );
    assert_eq!(out, vec!["15", "12", "24", "6", "1"]);
}

#[test]
fn bitwise_assign() {
    let out = run_js(
        r#"
let x = 0xFF;
x &= 0x0F;
console.log(x);
x |= 0xF0;
console.log(x);
x ^= 0xFF;
console.log(x);
"#,
    );
    assert_eq!(out, vec!["15", "255", "0"]);
}

#[test]
fn shift_assign() {
    let out = run_js(
        r#"
let x = 1;
x <<= 4;
console.log(x);
x >>= 2;
console.log(x);
"#,
    );
    assert_eq!(out, vec!["16", "4"]);
}

#[test]
fn bitwise_shift_count_is_masked_to_low_five_bits() {
    let out = run_js(
        r#"
console.log(-1 >> 33);
console.log(-16 >>> 33);
console.log(-16 >> 129);
"#,
    );
    assert_eq!(out, vec!["-1", "2147483640", "-8"]);
}

// ── Ternary ────────────────────────────────────────────────

#[test]
fn ternary_basic() {
    let out = run_js(
        r#"
const x = 5;
console.log(x > 3 ? "big" : "small");
console.log(x > 10 ? "big" : "small");
"#,
    );
    assert_eq!(out, vec!["big", "small"]);
}

#[test]
fn ternary_nested() {
    let out = run_js(
        r#"
const x = 5;
const result = x > 10 ? "big" : x > 3 ? "medium" : "small";
console.log(result);
"#,
    );
    assert_eq!(out, vec!["medium"]);
}

// ── typeof ─────────────────────────────────────────────────

#[test]
fn typeof_all_types() {
    let out = run_js(
        r#"
console.log(typeof 42);
console.log(typeof "hello");
console.log(typeof true);
console.log(typeof undefined);
console.log(typeof null);
console.log(typeof {});
console.log(typeof function(){});
"#,
    );
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
    let out = run_js(
        r#"
class Foo {}
const f = new Foo();
console.log(f instanceof Foo);
"#,
    );
    assert_eq!(out, vec!["true"]);
}

// ── in operator ────────────────────────────────────────────

#[test]
fn in_operator() {
    let out = run_js(
        r#"
const obj = { a: 1, b: 2 };
console.log("a" in obj);
console.log("c" in obj);
"#,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn in_operator_with_symbols() {
    let out = run_js(
        r#"
const key = Symbol("key");
const hidden = Symbol("hidden");
const obj = { [key]: "present" };

console.log(key in obj);
console.log(hidden in obj);
"#,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn in_operator_with_symbol_in_prototype_chain() {
    let out = run_js(
        r#"
const key = Symbol("token");
const proto = { [key]: "proto" };
const obj = Object.create(proto);
console.log(key in obj);
console.log(Object.getOwnPropertySymbols(obj).length);
"#,
    );
    assert_eq!(out, vec!["true", "0"]);
}

// ── delete ─────────────────────────────────────────────────

#[test]
fn delete_property() {
    let out = run_js(
        r#"
const obj = { a: 1, b: 2 };
delete obj.a;
console.log("a" in obj);
console.log(obj.b);
"#,
    );
    assert_eq!(out, vec!["false", "2"]);
}

#[test]
fn delete_non_configurable_property_returns_false() {
    let out = run_js(
        r#"
const obj = {};
Object.defineProperty(obj, "x", {
    value: 42,
    writable: false,
    configurable: false,
    enumerable: true,
});

console.log(delete obj.x);
console.log(obj.x);
"#,
    );
    assert_eq!(out, vec!["false", "42"]);
}

// ── void ───────────────────────────────────────────────────

#[test]
fn void_operator() {
    let out = run_js(
        r#"
console.log(void 0);
"#,
    );
    assert_eq!(out, vec!["undefined"]);
}

// ── Comma operator ─────────────────────────────────────────

#[test]
fn comma_operator() {
    let out = run_js(
        r#"
let x = (1, 2, 3);
console.log(x);
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn optional_chaining_computed_property() {
    let out = run_js(
        r#"
const obj = { nested: { value: 7 } };
const key = "nested";
console.log(obj?.[key]?.value);
console.log(obj?.["missing"]?.value);
"#,
    );
    assert_eq!(out, vec!["7", "undefined"]);
}

#[test]
fn bigints_mixed_with_numbers_throw_for_arithmetic() {
    let out = run_js(
        r#"
try {
    console.log(1n + 1);
} catch (e) {
    console.log(e.name);
}
console.log(String(8n / 3n));
console.log(String((-5n) % 3n));
"#,
    );
    assert_eq!(out, vec!["TypeError", "2", "-2"]);
}

#[test]
fn optional_chaining_computed_index() {
    let out = run_js(
        r#"
const list = [{ name: "a" }, { name: "b" }];
console.log(list?.[1]?.name);
console.log(list?.[5]?.name);
"#,
    );
    assert_eq!(out, vec!["b", "undefined"]);
}

#[test]
fn optional_chaining_call_skips_nonexistent_chain() {
    let out = run_js(
        r#"
let calls = 0;
const obj = {
    greet() {
        calls += 1;
        return "hello";
    }
};
console.log(obj?.greet?.());
console.log((null?.toString)?.());
console.log(calls);
"#,
    );
    assert_eq!(out, vec!["hello", "undefined", "1"]);
}

#[test]
fn optional_chaining_call_eager_property_lookup_only() {
    let out = run_js(
        r#"
let calls = 0;
const target = {
    get printer() {
        calls += 1;
        return () => {
            calls += 10;
            return calls;
        };
    }
};

console.log(target?.printer?.());
console.log((null?.printer)?.());
console.log(calls);
"#,
    );
    assert_eq!(out, vec!["11", "undefined", "11"]);
}

#[test]
fn typeof_undeclared_is_undefined() {
    let out = run_js(
        r#"
console.log(typeof definitelyMissing);
"#,
    );
    assert_eq!(out, vec!["undefined"]);
}

#[test]
fn comma_operator_evaluates_all_operands_left_to_right() {
    let out = run_js(
        r#"
const seen = [];
const value = (seen.push("a"), seen.push("b"), 5);
console.log(value);
console.log(seen.join(","));
"#,
    );
    assert_eq!(out, vec!["5", "a,b"]);
}

#[test]
fn modulo_sign_preserves_dividend_sign() {
    let out = run_js(
        r#"
console.log(10 % -3);
console.log(-10 % 3);
console.log(-10 % -3);
"#,
    );
    assert_eq!(out, vec!["1", "-1", "-1"]);
}

#[test]
fn nullish_coalescing_preserves_falsey_non_nullish_values() {
    let out = run_js(
        r#"
console.log(false ?? true);
console.log(0 ?? 10);
console.log("" ?? "fallback");
"#,
    );
    assert_eq!(out, vec!["false", "0", ""]);
}

#[test]
fn delete_array_element_keeps_length() {
    let out = run_js(
        r#"
const arr = [10, 20, 30];
delete arr[1];
console.log(arr.length);
console.log(1 in arr);
console.log(arr[1]);
"#,
    );
    assert_eq!(out, vec!["3", "false", "undefined"]);
}

#[test]
fn in_operator_on_array_indexes() {
    let out = run_js(
        r#"
const arr = ["x", "y"];
console.log(0 in arr);
console.log(2 in arr);
console.log("length" in arr);
"#,
    );
    assert_eq!(out, vec!["true", "false", "true"]);
}

#[test]
fn logical_or_returns_original_operand() {
    let out = run_js(
        r#"
console.log("left" || "right");
console.log(0 || 5);
"#,
    );
    assert_eq!(out, vec!["left", "5"]);
}

#[test]
fn logical_and_returns_original_operand() {
    let out = run_js(
        r#"
console.log("left" && "right");
console.log(0 && 5);
"#,
    );
    assert_eq!(out, vec!["right", "0"]);
}

#[test]
fn ternary_uses_truthiness() {
    let out = run_js(
        r#"
console.log("" ? "yes" : "no");
console.log([] ? "yes" : "no");
"#,
    );
    assert_eq!(out, vec!["no", "yes"]);
}

#[test]
fn typeof_array_and_class() {
    let out = run_js(
        r#"
class Box {}
console.log(typeof []);
console.log(typeof Box);
"#,
    );
    assert_eq!(out, vec!["object", "function"]);
}

#[test]
fn loose_equality_with_booleans_and_strings() {
    let out = run_js(
        r#"
console.log("0" == false);
console.log("1" == true);
console.log(2 == true);
"#,
    );
    assert_eq!(out, vec!["true", "true", "false"]);
}

#[test]
fn strict_equality_distinguishes_nan() {
    let out = run_js(
        r#"
console.log(NaN === NaN);
console.log(Object.is(NaN, NaN));
"#,
    );
    assert_eq!(out, vec!["false", "true"]);
}

#[test]
fn comparison_of_strings_is_lexicographic() {
    let out = run_js(
        r#"
console.log("2" > "10");
console.log("apple" < "banana");
"#,
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn exponentiation_is_right_associative() {
    let out = run_js(
        r#"
console.log(2 ** 3 * 2);
console.log(2 ** (3 + 1));
console.log(-(2 ** 3));
"#,
    );
    assert_eq!(out, vec!["16", "16", "-8"]);
}

#[test]
fn accessor_property_arithmetic_assignment_uses_getter_setter() {
    let out = run_js(
        r#"
const obj = {
    _x: 1,
    get x() {
        return this._x;
    },
    set x(v) {
        this._x = v;
    }
};

obj.x += 4;
console.log(obj.x);
console.log(obj._x);
"#,
    );
    assert_eq!(out, vec!["5", "5"]);
}

#[test]
fn delete_literal_expression_is_true() {
    let out = run_js(
        r#"
console.log(delete 123);
console.log(delete "x");
"#
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn typeof_bigint_and_symbol() {
    let out = run_js(
        r#"
console.log(typeof 100n);
console.log(typeof Symbol("id"));
"#,
    );
    assert_eq!(out, vec!["bigint", "symbol"]);
}

#[test]
fn prefix_and_postfix_increment_on_accessor_property() {
    let out = run_js(
        r#"
const obj = {
    _val: 10,
    get val() { return this._val; },
    set val(v) { this._val = v; }
};
console.log(++obj.val);
console.log(obj.val);
console.log(obj.val++);
console.log(obj.val);
"#,
    );
    assert_eq!(out, vec!["11", "11", "11", "12"]);
}

#[test]
fn nullish_assignment_short_circuit_eval() {
    let out = run_js(
        r#"
let sideEffect = false;
let val = "initial";
val ??= (sideEffect = true, "fallback");
console.log(val);
console.log(sideEffect);
"#,
    );
    assert_eq!(out, vec!["initial", "false"]);
}

