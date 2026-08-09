use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Relational Operators (`in`, `instanceof`, `<`, `>`, `<=`, `>=`)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_in_operator_own_and_inherited_properties() {
    let src = r#"
const proto = { parentKey: 1 };
const obj = Object.create(proto);
obj.ownKey = 2;

console.log(`${"ownKey" in obj}:${"parentKey" in obj}:${"missing" in obj}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:false"]);
}

#[test]
fn test_js_in_operator_array_indices() {
    let src = r#"
const arr = [10, , 30];
console.log(`${0 in arr}:${1 in arr}:${2 in arr}:${3 in arr}`);
"#;
    assert_eq!(run_js(src), vec!["true:false:true:false"]);
}

#[test]
fn test_js_in_operator_symbol_property() {
    let src = r#"
const sym = Symbol("key");
const obj = { [sym]: 42 };
console.log(sym in obj);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_in_operator_non_object_target_throws_typeerror() {
    let src = r#"
try {
    "length" in "string_primitive";
} catch (e) {
    console.log("in Operator Non-Object TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["in Operator Non-Object TypeError"]);
}

#[test]
fn test_js_instanceof_operator_prototype_chain() {
    let src = r#"
class Base {}
class Derived extends Base {}
const d = new Derived();

console.log(`${d instanceof Derived}:${d instanceof Base}:${d instanceof Object}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:true"]);
}

#[test]
fn test_js_instanceof_operator_builtin_objects() {
    let src = r#"
console.log(`${[] instanceof Array}:${({}) instanceof Object}:${(new Map()) instanceof Map}:${(new Date()) instanceof Date}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:true:true"]);
}

#[test]
fn test_js_instanceof_operator_primitive_lhs_returns_false() {
    let src = r#"
console.log(`${"hello" instanceof String}:${123 instanceof Number}:${true instanceof Boolean}`);
"#;
    assert_eq!(run_js(src), vec!["false:false:false"]); // Primitive LHS returns false!
}

#[test]
fn test_js_instanceof_operator_non_callable_rhs_throws_typeerror() {
    let src = r#"
try {
    ({}) instanceof {};
} catch (e) {
    console.log("instanceof Non-Callable RHS TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["instanceof Non-Callable RHS TypeError"]);
}

#[test]
fn test_js_instanceof_symbol_hasinstance_truthy_falsy_returns_boolean() {
    assert_eq!(
        run_js(
            r#"
class TruthyInstance {
    static [Symbol.hasInstance]() { return "truthy"; }
}
class FalsyInstance {
    static [Symbol.hasInstance]() { return ""; }
}
console.log({} instanceof TruthyInstance);
console.log({} instanceof FalsyInstance);
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn test_js_instanceof_symbol_hasinstance_observes_argument() {
    assert_eq!(
        run_js(
            r#"
const log = [];
class Tracker {
    static [Symbol.hasInstance](value) {
        log.push(typeof value);
        return value && value.token === "ok";
    }
}
console.log({ token: "ok" } instanceof Tracker);
console.log({ token: "bad" } instanceof Tracker);
console.log(log.join("|"));
"#
        ),
        vec!["true", "false", "object|object"]
    );
}

#[test]
fn test_js_in_operator_works_with_string_object_wrappers() {
    assert_eq!(
        run_js(
            r#"
console.log("0" in new String("abc"));
console.log(0 in new String("abc"));
console.log("length" in new String(""));
"#
        ),
        vec!["true", "true", "true"]
    );
}

#[test]
fn test_js_relational_numeric_comparisons() {
    let src = r#"
console.log(`${5 < 10}:${5 > 10}:${5 <= 5}:${5 >= 6}`);
"#;
    assert_eq!(run_js(src), vec!["true:false:true:false"]);
}

#[test]
fn test_js_relational_string_lexicographical_comparisons() {
    let src = r#"
console.log(`${"apple" < "banana"}:${"2" > "10"}:${"a" <= "a"}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:true"]); // "2" > "10" is true in lexicographical comparison!
}

#[test]
fn test_js_relational_mixed_number_and_string_comparison() {
    let src = r#"
console.log(`${2 > "10"}:${"2" < 10}:${"20" >= 20}`);
"#;
    assert_eq!(run_js(src), vec!["false:true:true"]); // Converts string to number when one operand is number!
}

#[test]
fn test_js_relational_nan_comparison_always_false() {
    let src = r#"
console.log(`${NaN < 5}:${NaN > 5}:${NaN <= 5}:${NaN >= 5}:${NaN < NaN}:${NaN <= NaN}`);
"#;
    assert_eq!(run_js(src), vec!["false:false:false:false:false:false"]);
}

#[test]
fn test_js_relational_infinity_comparisons() {
    let src = r#"
console.log(`${Infinity > 1e300}:${-Infinity < -1e300}:${Infinity >= Infinity}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:true"]);
}

#[test]
fn test_js_relational_bigint_comparisons() {
    let src = r#"
console.log(`${10n < 20n}:${100n >= 100n}:${-5n > -10n}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:true"]);
}

#[test]
fn test_js_relational_mixed_bigint_and_number_comparison() {
    let src = r#"
console.log(`${10n < 20}:${100n >= 100}:${5n > 10}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:false"]); // BigInt and Number can be compared with relational operators!
}

#[test]
fn test_js_relational_object_toprimitive_coercion() {
    let src = r#"
const obj1 = { [Symbol.toPrimitive]: () => 10 };
const obj2 = { valueOf: () => 20 };
console.log(obj1 < obj2);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_relational_null_and_undefined_coercion() {
    let src = r#"
console.log(`${null < 1}:${null >= 0}:${undefined < 1}:${undefined >= 0}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:false:false"]); // null is coerced to 0, undefined is coerced to NaN
}

#[test]
fn test_js_instanceof_custom_symbol_has_instance() {
    let src = r#"
const OddChecker = {
    [Symbol.hasInstance](val) {
        return typeof val === "number" && val % 2 !== 0;
    }
};
console.log((5 instanceof OddChecker) + "|" + (4 instanceof OddChecker));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_in_operator_private_field_brand_check() {
    let src = r#"
class Account {
    #secret = 123;
    static isAccount(obj) {
        return #secret in obj;
    }
}
const acc = new Account();
console.log(Account.isAccount(acc) + "|" + Account.isAccount({}));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_instanceof_rhs_null_or_undefined_returns_false() {
    assert_eq!(
        run_js(
            r#"
const obj = {};
try {
    console.log(obj instanceof null);
} catch (e) {
    console.log("null");
}
try {
    console.log(obj instanceof undefined);
} catch (e) {
    console.log("undefined");
}
"#
        ),
        vec!["false", "false"]
    );
}

#[test]
fn test_js_relational_symbol_operand_throws_typeerror() {
    let src = r#"
try {
    Symbol("a") < Symbol("b");
} catch (e) {
    console.log("Relational Symbol TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Relational Symbol TypeError"]);
}

#[test]
fn test_js_instanceof_uses_current_prototype_property_of_rhs() {
    let src = r#"
function C() {}
const before = new C();
console.log(before instanceof C);

C.prototype = {};
const after = new C();
console.log(after instanceof C);
console.log(before instanceof C);
"#;
    assert_eq!(run_js(src), vec!["true", "true", "false"]);
}

#[test]
fn test_js_instanceof_with_symbol_instanceof_primitive_wrapper() {
    let src = r#"
console.log(`foo` instanceof String);
console.log(new String("foo") instanceof String);
"#;
    assert_eq!(run_js(src), vec!["false", "true"]);
}

#[test]
fn test_js_in_operator_numeric_left_operand_coerces_to_property_key() {
    let src = r#"
const arr = ["a", "b", "c"];
console.log(`${1 in arr}:${"1" in arr}`);
console.log(`${99 in arr}:${99n in arr}`);
"#;
    assert_eq!(run_js(src), vec!["true:true", "false:true"]);
}

#[test]
fn test_js_instanceof_function_with_non_object_prototype_throws() {
    let src = r#"
function Weird() {}
Weird.prototype = 123;
console.log(({}) instanceof Weird);
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_relational_boolean_and_number_coercion() {
    let src = r#"
console.log(`${false < true}:${false <= 0}:${true > false}:${true >= 1}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:true:true"]);
}

#[test]
fn test_js_in_operator_integer_property_key_coercion() {
    assert_eq!(
        run_js(
            r#"
console.log(`${0 in { 0: "zero", 1: "one" }}:${2 in { 0: "zero", 1: "one" }}`);
"#
        ),
        vec!["true:false"]
    );
}

#[test]
fn test_js_instanceof_boolean_rhs_returns_false() {
    assert_eq!(run_js("console.log(({}) instanceof true);"), vec!["false"]);
}

#[test]
fn test_js_relational_evaluation_order_left_to_right_side_effects() {
    let src = r#"
const log = [];
const left = {
    [Symbol.toPrimitive]() {
        log.push("L");
        return 10;
    }
};
const right = {
    [Symbol.toPrimitive]() {
        log.push("R");
        return 5;
    }
};
console.log((left > right) + "|" + log.join(","));
"#;
    assert_eq!(run_js(src), vec!["true|L,R"]);
}

#[test]
fn test_js_in_operator_proxy_has_trap() {
    let src = r#"
const p = new Proxy({}, {
    has(target, prop) {
        return prop === "secret";
    }
});
console.log(`${"secret" in p}:${"other" in p}`);
"#;
    assert_eq!(run_js(src), vec!["true:false"]);
}
