/// JavaScript type coercion edge cases, modern features:
/// numeric separators, Error.cause, Object.hasOwn, globalThis,
/// structuredClone, Map/Set conversions, queueMicrotask,
/// string coercion, boolean coercion, numeric coercion deep patterns.
use super::helpers::run_js;

// ===================================================================
// TYPE COERCION: BOOLEAN
// ===================================================================

#[test]
fn coerce_falsy_values() {
    assert_eq!(
        run_js(
            r#"
let falsies = [false, 0, -0, "", null, undefined, NaN];
falsies.forEach(v => console.log(Boolean(v)));
"#
        ),
        &[
            "false", "false", "false", "false", "false", "false", "false"
        ]
    );
}

#[test]
fn coerce_truthy_values() {
    assert_eq!(
        run_js(
            r#"
let truthies = [true, 1, -1, "hello", {}, [], "false", "0"];
truthies.forEach(v => console.log(Boolean(v)));
"#
        ),
        &[
            "true", "true", "true", "true", "true", "true", "true", "true"
        ]
    );
}

#[test]
fn coerce_double_bang() {
    assert_eq!(
        run_js(
            r#"
console.log(!!0);
console.log(!!"");
console.log(!!null);
console.log(!!1);
console.log(!!"hi");
console.log(!!{});
"#
        ),
        &["false", "false", "false", "true", "true", "true"]
    );
}

// ===================================================================
// TYPE COERCION: NUMERIC
// ===================================================================

#[test]
fn coerce_string_to_number() {
    assert_eq!(
        run_js(
            r#"
console.log(Number("42"));
console.log(Number("3.14"));
console.log(Number(""));
console.log(Number(" "));
console.log(Number("hello"));
console.log(Number(true));
console.log(Number(false));
console.log(Number(null));
"#
        ),
        &["42", "3.14", "0", "0", "NaN", "1", "0", "0"]
    );
}

#[test]
fn coerce_unary_plus() {
    assert_eq!(
        run_js(
            r#"
console.log(+"42");
console.log(+"");
console.log(+true);
console.log(+false);
console.log(+null);
"#
        ),
        &["42", "0", "1", "0", "0"]
    );
}

#[test]
fn coerce_arithmetic_operators() {
    assert_eq!(
        run_js(
            r#"
console.log("5" - 2);
console.log("5" * 2);
console.log("5" / 2);
console.log("5" % 2);
console.log("5" + 2);
"#
        ),
        &["3", "10", "2.5", "1", "52"]
    );
}

// ===================================================================
// TYPE COERCION: STRING
// ===================================================================

#[test]
fn coerce_to_string() {
    assert_eq!(
        run_js(
            r#"
console.log(String(42));
console.log(String(true));
console.log(String(false));
console.log(String(null));
console.log(String(undefined));
console.log(String([]));
"#
        ),
        &["42", "true", "false", "null", "undefined", ""]
    );
}

#[test]
fn coerce_concat_with_plus() {
    assert_eq!(
        run_js(
            r#"
console.log("" + 42);
console.log("" + true);
console.log("" + null);
console.log("" + undefined);
console.log("" + [1,2,3]);
"#
        ),
        &["42", "true", "null", "undefined", "1,2,3"]
    );
}

// ===================================================================
// EQUALITY EDGE CASES
// ===================================================================

#[test]
fn loose_equality_quirks() {
    assert_eq!(
        run_js(
            r#"
console.log(null == undefined);
console.log(null == 0);
console.log(null == "");
console.log(null == false);
console.log("" == 0);
console.log("" == false);
console.log("0" == false);
console.log([] == false);
console.log([] == 0);
"#
        ),
        &[
            "true", "false", "false", "false", "true", "true", "true", "true", "true"
        ]
    );
}

#[test]
fn strict_equality_no_coercion() {
    assert_eq!(
        run_js(
            r#"
console.log(null === undefined);
console.log("" === 0);
console.log("0" === false);
console.log(0 === false);
console.log("" === false);
"#
        ),
        &["false", "false", "false", "false", "false"]
    );
}

// ===================================================================
// NUMERIC SEPARATORS
// ===================================================================

#[test]
fn numeric_separator() {
    assert_eq!(
        run_js(
            r#"
let million = 1_000_000;
let hex = 0xFF_FF;
let binary = 0b1010_0001;
console.log(million);
console.log(hex);
console.log(binary);
"#
        ),
        &["1000000", "65535", "161"]
    );
}

// ===================================================================
// ERROR.CAUSE
// ===================================================================

#[test]
fn error_cause() {
    assert_eq!(
        run_js(
            r#"
try {
    try {
        throw new Error("original");
    } catch (e) {
        throw new Error("wrapped", { cause: e });
    }
} catch (e) {
    console.log(e.message);
    console.log(e.cause.message);
}
"#
        ),
        &["wrapped", "original"]
    );
}

// ===================================================================
// OBJECT.HASOWN
// ===================================================================

#[test]
fn object_hasown() {
    assert_eq!(
        run_js(
            r#"
let obj = { a: 1 };
console.log(Object.hasOwn(obj, "a"));
console.log(Object.hasOwn(obj, "toString"));
"#
        ),
        &["true", "false"]
    );
}

#[test]
fn object_hasown_vs_in() {
    assert_eq!(
        run_js(
            r#"
let parent = { inherited: true };
let child = Object.create(parent);
child.own = true;
console.log("inherited" in child);
console.log(Object.hasOwn(child, "inherited"));
console.log(Object.hasOwn(child, "own"));
"#
        ),
        &["true", "false", "true"]
    );
}

// ===================================================================
// GLOBALTHIS
// ===================================================================

#[test]
fn globalthis_exists() {
    assert_eq!(
        run_js(
            r#"
console.log(typeof globalThis);
console.log(globalThis !== undefined);
"#
        ),
        &["object", "true"]
    );
}

// ===================================================================
// MAP / SET CONVERSIONS
// ===================================================================

#[test]
fn map_from_array() {
    assert_eq!(
        run_js(
            r#"
let m = new Map([["a", 1], ["b", 2], ["c", 3]]);
console.log(m.size);
console.log(m.get("b"));
"#
        ),
        &["3", "2"]
    );
}

#[test]
fn map_to_array() {
    assert_eq!(
        run_js(
            r#"
let m = new Map([["x", 10], ["y", 20]]);
let arr = Array.from(m);
console.log(arr.length);
console.log(arr[0][0] + "=" + arr[0][1]);
"#
        ),
        &["2", "x=10"]
    );
}

#[test]
fn map_to_object() {
    assert_eq!(
        run_js(
            r#"
let m = new Map([["a", 1], ["b", 2]]);
let obj = Object.fromEntries(m);
console.log(obj.a);
console.log(obj.b);
"#
        ),
        &["1", "2"]
    );
}

#[test]
fn object_to_map() {
    assert_eq!(
        run_js(
            r#"
let obj = { x: 10, y: 20 };
let m = new Map(Object.entries(obj));
console.log(m.get("x"));
console.log(m.get("y"));
console.log(m.size);
"#
        ),
        &["10", "20", "2"]
    );
}

#[test]
fn set_from_array() {
    assert_eq!(
        run_js(
            r#"
let s = new Set([1, 2, 2, 3, 3, 3]);
console.log(s.size);
let arr = Array.from(s);
console.log(arr.join(","));
"#
        ),
        &["3", "1,2,3"]
    );
}

#[test]
fn set_to_array_spread() {
    assert_eq!(
        run_js(
            r#"
let s = new Set(["a", "b", "c"]);
let arr = [...s];
console.log(arr.join(","));
"#
        ),
        &["a,b,c"]
    );
}

#[test]
fn set_operations_union_intersection() {
    assert_eq!(
        run_js(
            r#"
let a = new Set([1, 2, 3, 4]);
let b = new Set([3, 4, 5, 6]);
let union = new Set([...a, ...b]);
let intersection = new Set([...a].filter(x => b.has(x)));
let difference = new Set([...a].filter(x => !b.has(x)));
console.log([...union].sort().join(","));
console.log([...intersection].sort().join(","));
console.log([...difference].sort().join(","));
"#
        ),
        &["1,2,3,4,5,6", "3,4", "1,2"]
    );
}

// ===================================================================
// MAP ITERATION
// ===================================================================

#[test]
fn map_foreach() {
    assert_eq!(
        run_js(
            r#"
let m = new Map([["a", 1], ["b", 2], ["c", 3]]);
let result = [];
m.forEach((val, key) => result.push(key + "=" + val));
console.log(result.join(","));
"#
        ),
        &["a=1,b=2,c=3"]
    );
}

#[test]
fn map_for_of_destructure() {
    assert_eq!(
        run_js(
            r#"
let m = new Map([["x", 10], ["y", 20]]);
for (let [k, v] of m) {
    console.log(k + ":" + v);
}
"#
        ),
        &["x:10", "y:20"]
    );
}

// ===================================================================
// STRUCTUREDCLONE
// ===================================================================

#[test]
fn structuredclone_deep_copy() {
    assert_eq!(
        run_js(
            r#"
let orig = { a: 1, b: { c: 2, d: [3, 4] } };
let clone = structuredClone(orig);
clone.b.c = 99;
clone.b.d.push(5);
console.log(orig.b.c);
console.log(orig.b.d.length);
console.log(clone.b.c);
console.log(clone.b.d.length);
"#
        ),
        &["2", "2", "99", "3"]
    );
}

// ===================================================================
// MISCELLANEOUS
// ===================================================================

#[test]
fn string_raw_tag() {
    assert_eq!(
        run_js(
            r#"
let s = String.raw`Hello\nWorld`;
console.log(s);
console.log(s.includes("\\n"));
"#
        ),
        &["Hello\\nWorld", "true"]
    );
}

#[test]
fn array_isarray_edge_cases() {
    assert_eq!(
        run_js(
            r#"
console.log(Array.isArray([]));
console.log(Array.isArray(new Array()));
console.log(Array.isArray({}));
console.log(Array.isArray("string"));
console.log(Array.isArray(Array.of(1, 2)));
"#
        ),
        &["true", "true", "false", "false", "true"]
    );
}

#[test]
fn typeof_all_primitives() {
    assert_eq!(
        run_js(
            r#"
console.log(typeof undefined);
console.log(typeof null);
console.log(typeof true);
console.log(typeof 42);
console.log(typeof "hello");
console.log(typeof Symbol());
console.log(typeof function(){});
console.log(typeof {});
"#
        ),
        &[
            "undefined",
            "object",
            "boolean",
            "number",
            "string",
            "symbol",
            "function",
            "object"
        ]
    );
}

#[test]
fn optional_chaining_all_forms() {
    assert_eq!(
        run_js(
            r#"
let obj = {
    a: { b: { c: 42 } },
    fn: () => "called"
};
console.log(obj?.a?.b?.c);
console.log(obj?.x?.y?.z);
console.log(obj?.fn?.());
console.log(obj?.missing?.());
"#
        ),
        &["42", "undefined", "called", "undefined"]
    );
}

#[test]
fn number_coercion_of_arrays() {
    assert_eq!(
        run_js(
            r#"
console.log(Number([]));
console.log(Number([5]));
console.log(Number([1, 2]));
"#
        ),
        &["0", "5", "NaN"]
    );
}

#[test]
fn string_coercion_of_objects_uses_default_tag() {
    assert_eq!(
        run_js(
            r#"
console.log(String({}));
console.log(String({ a: 1 }));
"#
        ),
        &["[object Object]", "[object Object]"]
    );
}

#[test]
fn boolean_coercion_of_symbols_and_functions() {
    assert_eq!(
        run_js(
            r#"
console.log(Boolean(Symbol("x")));
console.log(Boolean(function() {}));
"#
        ),
        &["true", "true"]
    );
}

#[test]
fn arithmetic_with_null_and_undefined() {
    assert_eq!(
        run_js(
            r#"
console.log(null + 1);
console.log(undefined + 1);
console.log(null == 0);
"#
        ),
        &["1", "NaN", "false"]
    );
}

#[test]
fn object_is_distinguishes_negative_zero() {
    assert_eq!(
        run_js(
            r#"
console.log(Object.is(0, -0));
console.log(0 === -0);
"#
        ),
        &["false", "true"]
    );
}
