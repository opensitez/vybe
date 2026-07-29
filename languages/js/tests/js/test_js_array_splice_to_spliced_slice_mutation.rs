use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Array Splicing & Slicing (`splice`, `toSpliced` ES2023, `slice`)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_array_splice_remove_and_insert() {
    let src = r#"
const arr = ["a", "b", "c", "d"];
const removed = arr.splice(1, 2, "X", "Y");
console.log(arr.join(",") + "|removed=" + removed.join(","));
"#;
    assert_eq!(run_js(src), vec!["a,X,Y,d|removed=b,c"]);
}

#[test]
fn test_js_array_tospliced_immutable_es2023() {
    let src = r#"
const arr = [1, 2, 3, 4];
const spliced = arr.toSpliced(1, 2, 99);
console.log(arr.join(",") + "|" + spliced.join(",") + "|isDifferent=" + (arr !== spliced));
"#;
    assert_eq!(run_js(src), vec!["1,2,3,4|1,99,4|isDifferent=true"]);
}

#[test]
fn test_js_array_slice_shallow_copy_subset() {
    let src = r#"
const arr = ["a", "b", "c", "d", "e"];
const sliced = arr.slice(1, 4);
console.log(sliced.join(",") + "|orig=" + arr.join(","));
"#;
    assert_eq!(run_js(src), vec!["b,c,d|orig=a,b,c,d,e"]);
}

#[test]
fn test_js_array_slice_negative_indices() {
    let src = r#"
const arr = [10, 20, 30, 40, 50];
const sub = arr.slice(-3, -1);
console.log(sub.join(","));
"#;
    assert_eq!(run_js(src), vec!["30,40"]);
}

#[test]
fn test_js_array_splice_negative_start_index() {
    let src = r#"
const arr = [1, 2, 3, 4, 5];
const removed = arr.splice(-2, 1);
console.log(arr.join(",") + "|removed=" + removed.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,3,5|removed=4"]);
}

#[test]
fn test_js_array_tospliced_negative_indices() {
    let src = r#"
const arr = [1, 2, 3, 4];
const result = arr.toSpliced(-2, 1, 99);
console.log(result.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,99,4"]);
}

#[test]
fn test_js_array_splice_omit_delete_count_deletes_to_end() {
    let src = r#"
const arr = [10, 20, 30, 40];
const removed = arr.splice(2);
console.log(arr.join(",") + "|removed=" + removed.join(","));
"#;
    assert_eq!(run_js(src), vec!["10,20|removed=30,40"]);
}

#[test]
fn test_js_array_slice_omit_end_slices_to_end() {
    let src = r#"
const arr = [1, 2, 3, 4];
console.log(arr.slice(2).join(","));
"#;
    assert_eq!(run_js(src), vec!["3,4"]);
}

#[test]
fn test_js_array_splice_zero_delete_count_inserts_only() {
    let src = r#"
const arr = [1, 4];
const removed = arr.splice(1, 0, 2, 3);
console.log(arr.join(",") + "|removedLen=" + removed.length);
"#;
    assert_eq!(run_js(src), vec!["1,2,3,4|removedLen=0"]);
}

#[test]
fn test_js_array_splice_start_greater_than_length_clamped() {
    let src = r#"
const arr = [1, 2];
arr.splice(10, 2, 3);
console.log(arr.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,3"]);
}

#[test]
fn test_js_array_slice_start_greater_than_length_returns_empty() {
    let src = r#"
const arr = [1, 2];
console.log(arr.slice(10).length);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_array_splice_delete_count_greater_than_remaining_clamped() {
    let src = r#"
const arr = [1, 2, 3, 4];
const removed = arr.splice(1, 100);
console.log(arr.join(",") + "|removed=" + removed.join(","));
"#;
    assert_eq!(run_js(src), vec!["1|removed=2,3,4"]);
}

#[test]
fn test_js_array_slice_shallow_copy_preserves_object_references() {
    let src = r#"
const obj = { id: 1 };
const arr = [obj];
const sliced = arr.slice();
console.log(sliced[0] === obj);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_array_splice_sparse_array_holes() {
    let src = r#"
const sparse = [1, , 3, 4];
const removed = sparse.splice(1, 2);
console.log(sparse.join(",") + "|removedLen=" + removed.length + "|removedHole=" + !(0 in removed));
"#;
    assert_eq!(run_js(src), vec!["1,4|removedLen=2|removedHole=true"]);
}

#[test]
fn test_js_array_tospliced_copies_holes_as_undefined() {
    let src = r#"
const sparse = [1, , 3];
const result = sparse.toSpliced(0, 0);
console.log(result.length + "|" + result.map(x => String(x)).join(","));
"#;
    assert_eq!(run_js(src), vec!["3|1,undefined,3"]);
}

#[test]
fn test_js_array_slice_sparse_array_holes_preserved() {
    let src = r#"
const sparse = [1, , 3];
const sliced = sparse.slice();
console.log(sliced.length + "|hasHole=" + !(1 in sliced));
"#;
    assert_eq!(run_js(src), vec!["3|hasHole=true"]);
}

#[test]
fn test_js_array_splice_frozen_array_throws_in_strict() {
    let src = r#"
const frozen = Object.freeze([1, 2, 3]);
try {
    "use strict";
    frozen.splice(0, 1);
} catch (e) {
    console.log("Splice Frozen Array TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Splice Frozen Array TypeError"]);
}

#[test]
fn test_js_array_tospliced_on_frozen_array_succeeds() {
    let src = r#"
const frozen = Object.freeze([1, 2, 3]);
const result = frozen.toSpliced(1, 1);
console.log(result.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,3"]);
}

#[test]
fn test_js_array_splice_subclass_species() {
    let src = r#"
class CustomArray extends Array {}
const ca = new CustomArray(1, 2, 3);
const removed = ca.splice(0, 1);
console.log(ca.join(",") + "|removedIsCustom=" + (removed instanceof CustomArray));
"#;
    assert_eq!(run_js(src), vec!["2,3|removedIsCustom=true"]);
}

#[test]
fn test_js_array_slice_subclass_species() {
    let src = r#"
class CustomArray extends Array {}
const ca = new CustomArray(1, 2, 3);
const sliced = ca.slice(1);
console.log(sliced.join(",") + "|slicedIsCustom=" + (sliced instanceof CustomArray));
"#;
    assert_eq!(run_js(src), vec!["2,3|slicedIsCustom=true"]);
}

#[test]
fn test_js_array_tospliced_species_returns_base_array() {
    let src = r#"
class CustomArray extends Array {}
const ca = new CustomArray(1, 2, 3);
const spliced = ca.toSpliced(0, 1);
console.log(spliced.join(",") + "|isCustom=" + (spliced instanceof CustomArray));
"#;
    assert_eq!(run_js(src), vec!["2,3|isCustom=false"]);
}

