use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Strict Equality (`===`, `!==`), `Object.is` (SameValue) & SameValueZero Algorithms
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_strict_equality_different_types_return_false() {
    let src = r#"
console.log(`${"5" === 5}:${true === 1}:${null === undefined}:${[] === false}`);
"#;
    assert_eq!(run_js(src), vec!["false:false:false:false"]);
}

#[test]
fn test_js_strict_equality_nan_is_false() {
    let src = r#"
console.log(`${NaN === NaN}:${NaN !== NaN}`);
"#;
    assert_eq!(run_js(src), vec!["false:true"]);
}

#[test]
fn test_js_strict_equality_signed_zeros() {
    let src = r#"
console.log(`${+0 === -0}:${+0 !== -0}`);
"#;
    assert_eq!(run_js(src), vec!["true:false"]);
}

#[test]
fn test_js_object_is_same_value_algorithm_nan() {
    let src = r#"
console.log(Object.is(NaN, NaN));
"#;
    assert_eq!(run_js(src), vec!["true"]); // Object.is(NaN, NaN) is true!
}

#[test]
fn test_js_object_is_same_value_algorithm_signed_zeros() {
    let src = r#"
console.log(`${Object.is(+0, -0)}:${Object.is(+0, +0)}:${Object.is(-0, -0)}`);
"#;
    assert_eq!(run_js(src), vec!["false:true:true"]); // Object.is(+0, -0) is false!
}

#[test]
fn test_js_same_value_zero_used_in_map_and_set() {
    let src = r#"
const set = new Set();
set.add(+0);
set.add(-0);
set.add(NaN);
set.add(NaN);

console.log(`${set.size}:${set.has(-0)}:${set.has(NaN)}`);
"#;
    assert_eq!(run_js(src), vec!["2:true:true"]); // Set uses SameValueZero: treats +0 & -0 as equal, NaN & NaN as equal!
}

#[test]
fn test_js_same_value_zero_used_in_array_includes() {
    let src = r#"
const arr = [+0, NaN];
console.log(`${arr.includes(-0)}:${arr.includes(NaN)}`);
"#;
    assert_eq!(run_js(src), vec!["true:true"]);
}

#[test]
fn test_js_strict_equality_bigint_and_number() {
    let src = r#"
console.log(`${10n === 10}:${10n === 10n}`);
"#;
    assert_eq!(run_js(src), vec!["false:true"]);
}

#[test]
fn test_js_strict_equality_symbol_identity() {
    let src = r#"
const s1 = Symbol("key");
const s2 = Symbol("key");
console.log(`${s1 === s1}:${s1 === s2}`);
"#;
    assert_eq!(run_js(src), vec!["true:false"]);
}

#[test]
fn test_js_strict_equality_object_references() {
    let src = r#"
const o1 = { x: 1 };
const o2 = { x: 1 };
const o3 = o1;
console.log(`${o1 === o2}:${o1 === o3}`);
"#;
    assert_eq!(run_js(src), vec!["false:true"]);
}

#[test]
fn test_js_strict_inequality_operator() {
    let src = r#"
console.log(`${5 !== "5"}:${5 !== 5}:${null !== undefined}`);
"#;
    assert_eq!(run_js(src), vec!["true:false:true"]);
}

#[test]
fn test_js_object_is_primitive_comparisons() {
    let src = r#"
console.log(`${Object.is("a", "a")}:${Object.is(true, true)}:${Object.is(null, null)}:${Object.is(undefined, undefined)}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:true:true"]);
}

#[test]
fn test_js_strict_equality_boxed_primitives() {
    let src = r#"
console.log(`${new Number(7) === 7}:${new Number(7) === new Number(7)}:${Object.is(new Number(7), new Number(7))}`);
console.log(`${new Boolean(true) === true}:${new String("x") === "x"}`);
"#;
    assert_eq!(run_js(src), vec!["false:false:false", "false:false"]);
}

#[test]
fn test_js_object_is_object_references() {
    let src = r#"
const obj = {};
console.log(`${Object.is(obj, obj)}:${Object.is({}, {})}`);
"#;
    assert_eq!(run_js(src), vec!["true:false"]);
}

#[test]
fn test_js_strict_equality_function_references() {
    let src = r#"
function fn() {}
const f1 = fn;
const f2 = () => {};
console.log(`${f1 === fn}:${f1 === f2}`);
"#;
    assert_eq!(run_js(src), vec!["true:false"]);
}

#[test]
fn test_js_same_value_zero_map_key_lookup() {
    let src = r#"
const map = new Map();
map.set(+0, "zero");
map.set(NaN, "not_a_number");

console.log(`${map.get(-0)}:${map.get(NaN)}`);
"#;
    assert_eq!(run_js(src), vec!["zero:not_a_number"]);
}

#[test]
fn test_js_map_key_is_set_once_for_same_value_zero_pairs() {
    assert_eq!(
        run_js(
            r#"
const map = new Map();
map.set(+0, "plus");
map.set(-0, "minus");
map.set(NaN, "first");
map.set(NaN, "second");
console.log(`${map.size}:${map.get(0)}:${map.get(-0)}:${map.get(NaN)}`);
"#
        ),
        vec!["2:minus:minus:second"]
    );
}

#[test]
fn test_js_strict_equality_string_primitives_vs_string_objects() {
    let src = r#"
const strPrim = "hello";
const strObj = new String("hello");
console.log(`${strPrim === "hello"}:${strPrim === strObj}`);
"#;
    assert_eq!(run_js(src), vec!["true:false"]);
}

#[test]
fn test_js_strict_equality_boolean_primitives_vs_boolean_objects() {
    let src = r#"
const boolPrim = true;
const boolObj = new Boolean(true);
console.log(`${boolPrim === true}:${boolPrim === boolObj}`);
"#;
    assert_eq!(run_js(src), vec!["true:false"]);
}

#[test]
fn test_js_strict_equality_number_primitives_vs_number_objects() {
    let src = r#"
const numPrim = 42;
const numObj = new Number(42);
console.log(`${numPrim === 42}:${numPrim === numObj}`);
"#;
    assert_eq!(run_js(src), vec!["true:false"]);
}

#[test]
fn test_js_object_is_number_object_identity() {
    let src = r#"
const n1 = new Number(7);
const n2 = new Number(7);
console.log(`${n1 === n2}:${Object.is(n1, n2)}`);
"#;
    assert_eq!(run_js(src), vec!["false:false"]);
}

#[test]
fn test_js_strict_equality_signed_zero_with_object_is_false() {
    let src = r#"
console.log(`${Object.is(+0, -0)}:${+0 === -0}:${Object.is(-0, -0)}:${Object.is(+0, +0)}`);
"#;
    assert_eq!(run_js(src), vec!["false:true:true:true"]);
}

#[test]
fn test_js_object_is_symbol_for_global_registry() {
    let src = r#"
const s1 = Symbol.for("reg");
const s2 = Symbol.for("reg");
console.log(Object.is(s1, s2));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_same_value_zero_array_index_of_vs_includes() {
    let src = r#"
const arr = [NaN];
console.log(`${arr.indexOf(NaN)}:${arr.includes(NaN)}`); // indexOf uses === (returns -1), includes uses SameValueZero (returns true)!
"#;
    assert_eq!(run_js(src), vec!["-1:true"]);
}

#[test]
fn test_js_object_is_wrapper_object_reference_vs_primitive() {
    let src = r#"
const sym = Symbol("eq");
const symObj = Object(sym);
console.log(`${Object.is(sym, sym)}:${Object.is(sym, Symbol("eq"))}:${Object.is(sym, symObj)}`);
"#;
    assert_eq!(run_js(src), vec!["true:false:false"]);
}

#[test]
fn test_js_object_is_object_wrapper_distinct_instances() {
    let src = r#"
const n1 = new Number(5);
const n2 = new Number(5);
const b1 = new Boolean(false);
const b2 = new Boolean(false);
console.log(`${Object.is(n1, n2)}:${Object.is(b1, b2)}:${Object.is(n1, n1)}`);
"#;
    assert_eq!(run_js(src), vec!["false:false:true"]);
}

#[test]
fn test_js_object_identity_distinguishes_boxed_wrappers() {
    assert_eq!(
        run_js(
            r#"
const set = new Set();
const a = new Number(7);
const b = new Number(7);
set.add(a);
set.add(b);

const map = new Map();
const o1 = { x: 1 };
const o2 = { x: 1 };
map.set(o1, "first");
map.set(o2, "second");

console.log(`${set.size}:${map.size}`);
"#,
        ),
        vec!["2:2"]
    );
}

#[test]
fn test_js_strict_equality_bound_function_is_distinct_reference() {
    let src = r#"
function base() {}
const bound = base.bind(null);
console.log(`${base === bound}:${Object.is(base, bound)}:${bound === bound}`);
"#;
    assert_eq!(run_js(src), vec!["false:false:true"]);
}

#[test]
fn test_js_strict_equality_distinguishes_string_wrappers() {
    let src = r#"
console.log(`${new String("x") === new String("x")}:${Object.is(new String("x"), new String("x"))}:${Object.is("x", "x")}`);
"#;
    assert_eq!(run_js(src), vec!["false:false:true"]);
}

#[test]
fn test_js_object_is_distinguishes_wrapped_and_primitive_values() {
    let src = r#"
console.log(`${Object.is(new Number(7), 7)}:${Object.is(new String("x"), "x")}:${Object.is(new Boolean(false), false)}`);
"#;
    assert_eq!(run_js(src), vec!["false:false:false"]);
}

#[test]
fn test_js_object_is_bigint_comparisons() {
    let src = r#"
console.log(`${Object.is(100n, 100n)}:${Object.is(100n, 100)}:${Object.is(0n, -0n)}`);
"#;
    assert_eq!(run_js(src), vec!["true:false:true"]);
}

#[test]
fn test_js_same_value_zero_typed_array_includes() {
    let src = r#"
const floats = new Float64Array([+0, NaN]);
console.log(`${floats.includes(-0)}:${floats.includes(NaN)}`);
"#;
    assert_eq!(run_js(src), vec!["true:true"]);
}

