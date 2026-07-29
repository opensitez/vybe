use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Immutable Change Array by Copy (`toReversed`, `toSpliced`, `toSorted`, `with`) (ES2023)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_array_to_reversed_returns_new_array() {
    let src = r#"
const orig = [1, 2, 3];
const reversed = orig.toReversed();
console.log(orig.join(",") + "|" + reversed.join(",") + "|" + (orig !== reversed));
"#;
    assert_eq!(run_js(src), vec!["1,2,3|3,2,1|true"]);
}

#[test]
fn test_js_array_to_sorted_returns_new_array() {
    let src = r#"
const orig = [3, 1, 2];
const sorted = orig.toSorted();
console.log(orig.join(",") + "|" + sorted.join(",") + "|" + (orig !== sorted));
"#;
    assert_eq!(run_js(src), vec!["3,1,2|1,2,3|true"]);
}

#[test]
fn test_js_array_to_sorted_custom_comparator() {
    let src = r#"
const orig = [10, 5, 20];
const sorted = orig.toSorted((a, b) => b - a);
console.log(orig.join(",") + "|" + sorted.join(","));
"#;
    assert_eq!(run_js(src), vec!["10,5,20|20,10,5"]);
}

#[test]
fn test_js_array_to_spliced_insertion_deletion() {
    let src = r#"
const orig = ["a", "b", "e"];
const spliced = orig.toSpliced(2, 0, "c", "d");
console.log(orig.join(",") + "|" + spliced.join(","));
"#;
    assert_eq!(run_js(src), vec!["a,b,e|a,b,c,d,e"]);
}

#[test]
fn test_js_array_with_method_element_replacement() {
    let src = r#"
const orig = ["x", "y", "z"];
const updated = orig.with(1, "Y");
console.log(orig.join(",") + "|" + updated.join(",") + "|" + (orig !== updated));
"#;
    assert_eq!(run_js(src), vec!["x,y,z|x,Y,z|true"]);
}

#[test]
fn test_js_array_with_negative_index() {
    let src = r#"
const orig = [10, 20, 30];
const updated = orig.with(-1, 99);
console.log(updated.join(","));
"#;
    assert_eq!(run_js(src), vec!["10,20,99"]);
}

#[test]
fn test_js_array_with_index_out_of_bounds_throws_rangeerror() {
    let src = r#"
const orig = [1, 2];
try {
    orig.with(5, 100);
} catch (e) {
    console.log("Array with Index Out of Bounds RangeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Array with Index Out of Bounds RangeError"]
    );
}

#[test]
fn test_js_array_with_negative_index_out_of_bounds_throws_rangeerror() {
    let src = r#"
const orig = [1, 2];
try {
    orig.with(-5, 100);
} catch (e) {
    console.log("Array with Negative Index Out of Bounds RangeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Array with Negative Index Out of Bounds RangeError"]
    );
}

#[test]
fn test_js_array_to_reversed_sparse_array_dense_copy() {
    let src = r#"
const sparse = [1, , 3];
const reversed = sparse.toReversed();
console.log(reversed.length + "|hasHole=" + !(1 in reversed) + "|val=" + reversed[1]);
"#;
    assert_eq!(run_js(src), vec!["3|hasHole=false|val=undefined"]); // Copy methods produce dense arrays!
}

#[test]
fn test_js_array_to_sorted_sparse_array_holes_moved_to_end() {
    let src = r#"
const sparse = [2, , 1];
const sorted = sparse.toSorted();
console.log(sorted.join(",") + "|len=" + sorted.length);
"#;
    assert_eq!(run_js(src), vec!["1,2,|len=3"]);
}

#[test]
fn test_js_typed_array_to_reversed() {
    let src = r#"
const u8 = new Uint8Array([1, 2, 3]);
const reversed = u8.toReversed();
console.log((reversed instanceof Uint8Array) + "|" + reversed.join(","));
"#;
    assert_eq!(run_js(src), vec!["true|3,2,1"]);
}

#[test]
fn test_js_typed_array_to_sorted() {
    let src = r#"
const u8 = new Uint8Array([30, 10, 20]);
const sorted = u8.toSorted();
console.log((sorted instanceof Uint8Array) + "|" + sorted.join(","));
"#;
    assert_eq!(run_js(src), vec!["true|10,20,30"]);
}

#[test]
fn test_js_typed_array_with_method() {
    let src = r#"
const u8 = new Uint8Array([10, 20, 30]);
const updated = u8.with(1, 99);
console.log((updated instanceof Uint8Array) + "|" + updated.join(","));
"#;
    assert_eq!(run_js(src), vec!["true|10,99,30"]);
}

#[test]
fn test_js_array_to_spliced_deletion_only() {
    let src = r#"
const orig = [1, 2, 3, 4, 5];
const spliced = orig.toSpliced(1, 3);
console.log(spliced.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,5"]);
}

#[test]
fn test_js_array_to_spliced_negative_start_index() {
    let src = r#"
const orig = ["a", "b", "c", "d"];
const spliced = orig.toSpliced(-2, 1, "X");
console.log(spliced.join(","));
"#;
    assert_eq!(run_js(src), vec!["a,b,X,d"]);
}

#[test]
fn test_js_array_to_spliced_negative_delete_count_clamped_to_zero() {
    let src = r#"
const orig = [1, 2, 3];
const spliced = orig.toSpliced(1, -2, 99);
console.log(spliced.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,99,2,3"]);
}

#[test]
fn test_js_array_to_reversed_array_like_object() {
    let src = r#"
const arrayLike = { 0: "a", 1: "b", length: 2 };
const reversed = Array.prototype.toReversed.call(arrayLike);
console.log(Array.isArray(reversed) + "|" + reversed.join(","));
"#;
    assert_eq!(run_js(src), vec!["true|b,a"]);
}

#[test]
fn test_js_array_to_sorted_array_like_object() {
    let src = r#"
const arrayLike = { 0: "z", 1: "a", length: 2 };
const sorted = Array.prototype.toSorted.call(arrayLike);
console.log(Array.isArray(sorted) + "|" + sorted.join(","));
"#;
    assert_eq!(run_js(src), vec!["true|a,z"]);
}

#[test]
fn test_js_array_with_array_like_object() {
    let src = r#"
const arrayLike = { 0: 10, 1: 20, length: 2 };
const updated = Array.prototype.with.call(arrayLike, 0, 99);
console.log(Array.isArray(updated) + "|" + updated.join(","));
"#;
    assert_eq!(run_js(src), vec!["true|99,20"]);
}

#[test]
fn test_js_array_to_sorted_non_callable_comparator_throws_typeerror() {
    let src = r#"
try {
    [1, 2].toSorted("not_a_function");
} catch (e) {
    console.log("toSorted Comparator TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["toSorted Comparator TypeError"]);
}

#[test]
fn test_js_array_toreversed_species_returns_base_array() {
    let src = r#"
class CustomArray extends Array {}
const ca = new CustomArray(1, 2, 3);
const rev = ca.toReversed();
console.log(rev.join(",") + "|isCustom=" + (rev instanceof CustomArray));
"#;
    assert_eq!(run_js(src), vec!["3,2,1|isCustom=false"]);
}

