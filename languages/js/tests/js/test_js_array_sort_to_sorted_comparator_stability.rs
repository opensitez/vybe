use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Array Sorting (`sort` In-Place, `toSorted` Immutable ES2023, Stability)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_array_sort_default_string_unicode_order() {
    let src = r#"
const items = [10, 2, 5, 1];
items.sort();
console.log(items.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,10,2,5"]); // Lexicographical string conversion ordering
}

#[test]
fn test_js_array_sort_custom_numeric_comparator() {
    let src = r#"
const nums = [10, 2, 5, 1];
nums.sort((a, b) => a - b);
console.log(nums.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,5,10"]);
}

#[test]
fn test_js_array_tosorted_immutable_es2023() {
    let src = r#"
const original = [3, 1, 2];
const sorted = original.toSorted((a, b) => a - b);
console.log(original.join(",") + "|" + sorted.join(",") + "|isDifferent=" + (original !== sorted));
"#;
    assert_eq!(run_js(src), vec!["3,1,2|1,2,3|isDifferent=true"]);
}

#[test]
fn test_js_array_sort_in_place_mutation() {
    let src = r#"
const arr = ["c", "a", "b"];
const res = arr.sort();
console.log(arr.join(",") + "|isSameRef=" + (res === arr));
"#;
    assert_eq!(run_js(src), vec!["a,b,c|isSameRef=true"]);
}

#[test]
fn test_js_array_sort_stability_es2019() {
    let src = r#"
const items = [
    { name: "A", score: 10 },
    { name: "B", score: 5 },
    { name: "C", score: 10 },
    { name: "D", score: 5 }
];
items.sort((a, b) => a.score - b.score);
console.log(items.map(i => i.name).join(",")); // Stable sort preserves relative order: B, D, A, C
"#;
    assert_eq!(run_js(src), vec!["B,D,A,C"]);
}

#[test]
fn test_js_array_sort_undefined_values_moved_to_end() {
    let src = r#"
const arr = [3, undefined, 1, undefined, 2];
arr.sort((a, b) => a - b);
console.log(arr.map(x => String(x)).join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,3,undefined,undefined"]);
}

#[test]
fn test_js_array_sort_sparse_holes_moved_to_end_after_undefined() {
    let src = r#"
const sparse = [2, , 1, undefined, 3];
sparse.sort((a, b) => a - b);
console.log(sparse.length + "|" + sparse.map(x => String(x)).join(","));
"#;
    assert_eq!(run_js(src), vec!["5|1,2,3,undefined,undefined"]);
}

#[test]
fn test_js_array_tosorted_copies_holes_as_undefined() {
    let src = r#"
const sparse = [2, , 1];
const sorted = sparse.toSorted((a, b) => a - b);
console.log(sorted.length + "|" + sorted.map(x => String(x)).join(","));
"#;
    assert_eq!(run_js(src), vec!["3|1,2,undefined"]);
}

#[test]
fn test_js_array_sort_comparator_returns_nan_or_zero() {
    let src = r#"
const nums = [3, 1, 2];
nums.sort(() => NaN); // Treated as 0 -> elements remain in place
console.log(nums.join(","));
"#;
    assert_eq!(run_js(src), vec!["3,1,2"]);
}

#[test]
fn test_js_array_sort_object_by_property() {
    let src = r#"
const users = [{ age: 30 }, { age: 20 }, { age: 25 }];
users.sort((a, b) => a.age - b.age);
console.log(users.map(u => u.age).join(","));
"#;
    assert_eq!(run_js(src), vec!["20,25,30"]);
}

#[test]
fn test_js_array_sort_reverse_comparator() {
    let src = r#"
const nums = [1, 2, 3, 4];
nums.sort((a, b) => b - a);
console.log(nums.join(","));
"#;
    assert_eq!(run_js(src), vec!["4,3,2,1"]);
}

#[test]
fn test_js_array_sort_non_callable_comparator_throws_typeerror() {
    let src = r#"
try {
    [1, 2].sort(123);
} catch (e) {
    console.log("Sort Invalid Comparator TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Sort Invalid Comparator TypeError"]);
}

#[test]
fn test_js_array_tosorted_non_callable_comparator_throws_typeerror() {
    let src = r#"
try {
    [1, 2].toSorted("not_fn");
} catch (e) {
    console.log("toSorted Invalid Comparator TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["toSorted Invalid Comparator TypeError"]);
}

#[test]
fn test_js_array_sort_null_comparator_throws_typeerror() {
    let src = r#"
try {
    [1, 2].sort(null);
} catch (e) {
    console.log("Sort Null Comparator TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Sort Null Comparator TypeError"]);
}

#[test]
fn test_js_array_sort_symbol_elements_throws_without_comparator() {
    let src = r#"
const symbols = [Symbol("b"), Symbol("a")];
try {
    symbols.sort(); // String conversion on Symbol throws TypeError!
} catch (e) {
    console.log("Sort Symbol Without Comparator TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Sort Symbol Without Comparator TypeError"]
    );
}

#[test]
fn test_js_array_sort_symbol_elements_with_custom_comparator() {
    let src = r#"
const s1 = Symbol("a"), s2 = Symbol("b");
const symbols = [s2, s1];
symbols.sort((a, b) => a.description.localeCompare(b.description));
console.log(symbols.map(s => s.description).join(","));
"#;
    assert_eq!(run_js(src), vec!["a,b"]);
}

#[test]
fn test_js_array_sort_string_locale_compare() {
    let src = r#"
const words = ["réservé", "premier", "cliché", "adieu"];
words.sort((a, b) => a.localeCompare(b));
console.log(words.join(","));
"#;
    assert_eq!(run_js(src), vec!["adieu,cliché,premier,réservé"]);
}

#[test]
fn test_js_array_sort_frozen_array_throws_in_strict() {
    let src = r#"
const frozen = Object.freeze([3, 1, 2]);
try {
    "use strict";
    frozen.sort();
} catch (e) {
    console.log("Sort Frozen Array TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Sort Frozen Array TypeError"]);
}

#[test]
fn test_js_array_tosorted_on_frozen_array_succeeds() {
    let src = r#"
const frozen = Object.freeze([3, 1, 2]);
const sorted = frozen.toSorted();
console.log(sorted.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,3"]);
}

#[test]
fn test_js_array_sort_subclass_species() {
    let src = r#"
class CustomArray extends Array {}
const ca = new CustomArray(3, 1, 2);
ca.sort();
console.log(ca.join(",") + "|isCustom=" + (ca instanceof CustomArray));
"#;
    assert_eq!(run_js(src), vec!["1,2,3|isCustom=true"]);
}

#[test]
fn test_js_array_sort_fractional_comparator_values() {
    let src = r#"
const nums = [3, 1, 2];
nums.sort((a, b) => a > b ? 0.5 : -0.5);
console.log(nums.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,3"]);
}
