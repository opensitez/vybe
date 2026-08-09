/// Type coercion deep — abstract equality, comparison operators, ToNumber,
/// ToString, ToPrimitive with hint, implicit conversions in operators,
/// valueOf/toString interaction, NaN propagation, symbol coercion.
use super::helpers::run_js;

// ── abstract equality (==) ────────────────────────────────────────────────────

#[test]
fn abstract_eq_null_undefined() {
    assert_eq!(
        run_js(
            r#"
console.log(null == undefined);
console.log(undefined == null);
console.log(null == 0);
console.log(undefined == 0);
console.log(null == false);
"#
        ),
        vec!["true", "true", "false", "false", "false"]
    );
}

#[test]
fn abstract_eq_number_string() {
    assert_eq!(
        run_js(
            r#"
console.log(1 == "1");
console.log(0 == "");
console.log(0 == "0");
"#
        ),
        vec!["true", "true", "true"]
    );
}

#[test]
fn abstract_eq_boolean_coerces_to_number() {
    assert_eq!(
        run_js(
            r#"
console.log(true == 1);
console.log(false == 0);
console.log(true == "1");
console.log(false == "");
"#
        ),
        vec!["true", "true", "true", "true"]
    );
}

#[test]
fn abstract_eq_object_to_primitive() {
    assert_eq!(
        run_js(
            r#"
const obj = { valueOf() { return 42; } };
console.log(obj == 42);
console.log(obj == "42");
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn abstract_eq_bigint_and_number_string() {
    assert_eq!(
        run_js(
            r#"
console.log(1n == 1);
console.log(1n == "1");
console.log(0n == false);
console.log(2n == 2.5);
console.log(1n === 1);
"#
        ),
        vec!["true", "true", "true", "false", "false"]
    );
}

#[test]
fn to_primitive_invalid_symbol_to_primitive_return_throws() {
    assert_eq!(
        run_js(
            r#"
const bad = {
    [Symbol.toPrimitive]() {
        return {};
    }
};

try {
    console.log(bad == 1);
} catch (e) {
    console.log(e.name);
}

console.log(bad + "x");
"#
        ),
        vec!["false", "[object]x"]
    );
}

// ── ToNumber rules ────────────────────────────────────────────────────────────

#[test]
fn to_number_strings() {
    assert_eq!(
        run_js(
            r#"
console.log(+"42");
console.log(+"3.14");
console.log(+"");
console.log(+"  ");
console.log(+"\t1\n");
"#
        ),
        vec!["42", "3.14", "0", "0", "1"]
    );
}

#[test]
fn to_number_non_numeric_string_is_nan() {
    assert_eq!(
        run_js(
            r#"
console.log(isNaN(+"hello"));
console.log(isNaN(+"1a"));
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn to_number_booleans_and_null() {
    assert_eq!(
        run_js(
            r#"
console.log(+true);
console.log(+false);
console.log(+null);
console.log(+undefined);
"#
        ),
        vec!["1", "0", "0", "NaN"]
    );
}

#[test]
fn to_number_symbol_throws_typeerror() {
    assert_eq!(
        run_js(
            r#"
try {
    console.log(+Symbol("token"));
} catch (e) {
    console.log(e.name);
}
"#,
        ),
        vec!["TypeError"]
    );
}

#[test]
fn to_number_array_single_element() {
    assert_eq!(
        run_js(
            r#"
console.log(+[]);
console.log(+[42]);
console.log(isNaN(+[1,2]));
"#
        ),
        vec!["0", "42", "true"]
    );
}

#[test]
fn symbol_in_string_interpolation_throws_without_explicit_tostring() {
    assert_eq!(
        run_js(
            r#"
const sym = Symbol("id");
try {
    console.log(`${sym}`);
} catch (e) {
    console.log(e.name);
}
console.log(String(sym));
"#,
        ),
        vec!["TypeError", "Symbol(id)"]
    );
}

// ── ToString coercion ─────────────────────────────────────────────────────────

#[test]
fn tostring_null_undefined() {
    assert_eq!(
        run_js(
            r#"
console.log(String(null));
console.log(String(undefined));
console.log("" + null);
console.log("" + undefined);
"#
        ),
        vec!["null", "undefined", "null", "undefined"]
    );
}

#[test]
fn tostring_numbers() {
    assert_eq!(
        run_js(
            r#"
console.log(String(0));
console.log(String(-0));
console.log(String(Infinity));
console.log(String(-Infinity));
console.log(String(NaN));
"#
        ),
        vec!["0", "0", "Infinity", "-Infinity", "NaN"]
    );
}

#[test]
fn tostring_object_calls_tostring_method() {
    assert_eq!(
        run_js(
            r#"
const obj = { toString() { return "custom"; } };
console.log("" + obj);
console.log(String(obj));
"#
        ),
        vec!["custom", "custom"]
    );
}

#[test]
fn tostring_bigint_and_number_from_explicit_calls() {
    assert_eq!(
        run_js(
            r#"
console.log(String(10n));
console.log(Number(10n));
console.log(Number(1n + 2n));
"#,
        ),
        vec!["10", "10", "3"]
    );
}

// ── ToPrimitive with hint ─────────────────────────────────────────────────────

#[test]
fn toprimitive_number_hint_prefers_valueof() {
    assert_eq!(
        run_js(
            r#"
const obj = {
    valueOf() { return 42; },
    toString() { return "str"; }
};
console.log(obj - 0);    // number hint → valueOf
console.log(`${obj}`);  // string hint → toString
"#
        ),
        vec!["42", "str"]
    );
}

#[test]
fn toprimitive_symbol_overrides_all() {
    assert_eq!(
        run_js(
            r#"
const obj = {
    [Symbol.toPrimitive](hint) {
        if (hint === "number") return 10;
        if (hint === "string") return "ten";
        return true; // default
    }
};
console.log(+obj);
console.log(`${obj}`);
console.log(obj + "");
"#
        ),
        vec!["10", "ten", "true"]
    );
}

#[test]
fn toprimitive_uses_valueof_first_and_can_still_use_tostring_when_required() {
    assert_eq!(
        run_js(
            r#"
const obj = {
    valueOf() {
        throw new Error("valueOf exploded");
    },
    toString() {
        return "stringified";
    }
};
try {
    console.log(+obj);
} catch (e) {
    console.log(e.message);
}
try {
    console.log(Number(obj));
} catch (e) {
    console.log(e.message);
}
console.log(String(obj));
"#
        ),
        vec!["valueOf exploded", "valueOf exploded", "stringified"]
    );
}

// ── addition operator ambiguity ───────────────────────────────────────────────

#[test]
fn addition_string_vs_number() {
    assert_eq!(
        run_js(
            r#"
console.log(1 + 2);
console.log("1" + 2);
console.log(1 + "2");
console.log("" + 1 + 2);
console.log(1 + 2 + "3");
"#
        ),
        vec!["3", "12", "12", "12", "33"]
    );
}

#[test]
fn addition_object_and_number() {
    assert_eq!(
        run_js(
            r#"
const obj = { valueOf() { return 5; } };
console.log(obj + 3);
console.log(3 + obj);
"#
        ),
        vec!["8", "8"]
    );
}

// ── NaN propagation ───────────────────────────────────────────────────────────

#[test]
fn nan_propagates_in_math() {
    assert_eq!(
        run_js(
            r#"
console.log(NaN + 1);
console.log(NaN * 5);
console.log(NaN - NaN);
console.log(0 / 0);
"#
        ),
        vec!["NaN", "NaN", "NaN", "NaN"]
    );
}

#[test]
fn nan_is_not_equal_to_itself() {
    assert_eq!(
        run_js(
            r#"
console.log(NaN === NaN);
console.log(NaN == NaN);
console.log(Number.isNaN(NaN));
"#
        ),
        vec!["false", "false", "true"]
    );
}

#[test]
fn to_boolean_truthy_falsy() {
    assert_eq!(
        run_js(
            r#"
console.log(Boolean(""));
console.log(Boolean("0"));
console.log(Boolean(0));
console.log(Boolean(NaN));
console.log(Boolean({}));
console.log(Boolean([]));
console.log(Boolean(Symbol("coercion")));
"#
        ),
        vec!["false", "true", "false", "false", "true", "true", "true"]
    );
}

#[test]
fn to_primitive_all_converters_non_primitive_throws_typeerror() {
    assert_eq!(
        run_js(
            r#"
const obj = {
    valueOf() { return {}; },
    toString() { return {}; }
};
console.log(obj + 1);
"#
        ),
        vec!["[object Object]1"]
    );
}

// ── comparison operators ──────────────────────────────────────────────────────

#[test]
fn comparison_string_vs_string_lexicographic() {
    assert_eq!(
        run_js(
            r#"
console.log("apple" < "banana");
console.log("10" < "9"); // string comparison, not numeric
console.log(10 < 9);     // numeric
"#
        ),
        vec!["true", "true", "false"]
    );
}

#[test]
fn comparison_with_null_and_zero() {
    assert_eq!(
        run_js(
            r#"
// These are counterintuitive
console.log(null > 0);
console.log(null == 0);
console.log(null >= 0);
"#
        ),
        vec!["false", "false", "true"]
    );
}

// ── symbol coercion ───────────────────────────────────────────────────────────

#[test]
fn symbol_to_number_throws() {
    assert_eq!(
        run_js(
            r#"
let threw = false;
try { +Symbol("test"); } catch (e) { threw = e instanceof TypeError; }
console.log(threw);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn symbol_to_string_via_template_throws() {
    assert_eq!(
        run_js(
            r#"
let threw = false;
try { `${Symbol("test")}`; } catch (e) { threw = e instanceof TypeError; }
console.log(threw);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn symbol_explicit_string_conversion_works() {
    assert_eq!(
        run_js(
            r#"
const sym = Symbol("desc");
console.log(String(sym));
console.log(sym.toString());
"#
        ),
        vec!["Symbol(desc)", "Symbol(desc)"]
    );
}

#[test]
fn relational_comparisons_follow_coercion_rules() {
    assert_eq!(
        run_js(
            r#"
console.log([1] > 0);       // array -> string -> number
console.log([1, 2] > 0);    // array with >1 element coerces to NaN
console.log(true > 0);
console.log(false < 1);
"#
        ),
        vec!["true", "false", "true", "true"]
    );
}

#[test]
fn comparison_with_undefined_is_always_false() {
    assert_eq!(
        run_js(
            r#"
console.log([undefined > 0, undefined < 0, undefined == 0].join("|"));
"#
        ),
        vec!["false|false|false"]
    );
}
