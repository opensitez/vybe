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
