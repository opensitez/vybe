use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Array Flattening (`flat`, `flatMap`, Depth Recursion)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_array_flat_default_depth_one() {
    let src = r#"
const nested = [1, [2, 3], [4, [5]]];
const flattened = nested.flat();
console.log(flattened.map(x => Array.isArray(x) ? x.join(":") : x).join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,3,4,5"]);
}

#[test]
fn test_js_array_flat_custom_depth() {
    let src = r#"
const nested = [1, [2, [3, [4]]]];
const f2 = nested.flat(2);
console.log(f2.map(x => Array.isArray(x) ? x.join(":") : x).join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,3,4"]);
}

#[test]
fn test_js_array_flat_infinity_depth() {
    let src = r#"
const deeplyNested = [1, [2, [3, [4, [5]]]]];
const flatAll = deeplyNested.flat(Infinity);
console.log(flatAll.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,3,4,5"]);
}

#[test]
fn test_js_array_flat_removes_sparse_array_holes() {
    let src = r#"
const sparse = [1, , 3, [4, , 5]];
const flatSparse = sparse.flat();
console.log(flatSparse.length + "|" + flatSparse.join(","));
"#;
    assert_eq!(run_js(src), vec!["4|1,3,4,5"]);
}

#[test]
fn test_js_array_flatmap_mapping_and_flattening() {
    let src = r#"
const nums = [1, 2, 3];
const res = nums.flatMap(x => [x, x * 2]);
console.log(res.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,2,4,3,6"]);
}

#[test]
fn test_js_array_flatmap_depth_always_one() {
    let src = r#"
const nums = [1, 2];
const res = nums.flatMap(x => [[x * 10]]);
console.log(res.map(x => x.join(":")).join(","));
"#;
    assert_eq!(run_js(src), vec!["10,20"]);
}

#[test]
fn test_js_array_flatmap_filtering_pattern() {
    let src = r#"
const sentence = ["Hello World", "JS Engine"];
const words = sentence.flatMap(s => s.split(" "));
console.log(words.join(","));
"#;
    assert_eq!(run_js(src), vec!["Hello,World,JS,Engine"]);
}

#[test]
fn test_js_array_flatmap_drop_elements_returns_empty_array() {
    let src = r#"
const nums = [1, 2, 3, 4];
const evensOnly = nums.flatMap(x => x % 2 === 0 ? [x] : []);
console.log(evensOnly.join(","));
"#;
    assert_eq!(run_js(src), vec!["2,4"]);
}

#[test]
fn test_js_array_flat_zero_depth_returns_shallow_copy() {
    let src = r#"
const arr = [1, [2]];
const f0 = arr.flat(0);
console.log(f0.length + "|" + (f0 !== arr));
"#;
    assert_eq!(run_js(src), vec!["2|true"]);
}

#[test]
fn test_js_array_flat_negative_depth_treated_as_zero() {
    let src = r#"
const arr = [1, [2]];
const fNeg = arr.flat(-5);
console.log(fNeg.length);
"#;
    assert_eq!(run_js(src), vec!["2"]);
}

#[test]
fn test_js_array_flatmap_this_argument_binding() {
    let src = r#"
const ctx = { factor: 100 };
const nums = [1, 2];
const res = nums.flatMap(function(x) { return [x * this.factor]; }, ctx);
console.log(res.join(","));
"#;
    assert_eq!(run_js(src), vec!["100,200"]);
}

#[test]
fn test_js_array_flatmap_index_and_array_arguments() {
    let src = r#"
const items = ["a", "b"];
const res = items.flatMap((val, idx) => [`${val}${idx}`]);
console.log(res.join(","));
"#;
    assert_eq!(run_js(src), vec!["a0,b1"]);
}

#[test]
fn test_js_array_flat_non_array_elements_not_flattened() {
    let src = r#"
const arrayLikeObject = { 0: "a", 1: "b", length: 2 };
const arr = [1, arrayLikeObject];
const flattened = arr.flat();
console.log(flattened.length + "|" + (typeof flattened[1]));
"#;
    assert_eq!(run_js(src), vec!["2|object"]);
}

#[test]
fn test_js_array_flatmap_non_callable_callback_throws() {
    let src = r#"
try {
    [1, 2].flatMap("not_callable");
} catch (e) {
    console.log("flatMap Non-Callable TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["flatMap Non-Callable TypeError"]);
}

#[test]
fn test_js_array_flat_string_depth_coercion() {
    let src = r#"
const nested = [1, [2, [3]]];
const flattened = nested.flat("2");
console.log(flattened.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,3"]);
}

#[test]
fn test_js_array_flatmap_returning_non_array_iterable() {
    let src = r#"
const set = new Set([10, 20]);
const res = [1].flatMap(() => set);
console.log(res.join(",") + "|isArr=" + Array.isArray(res)); // Returns flattened Array, not Set!
"#;
    assert_eq!(run_js(src), vec!["10,20|isArr=true"]);
}

#[test]
fn test_js_array_flat_circular_reference_recursion_limit() {
    let src = r#"
const arr = [1];
arr.push(arr); // Circular nested array
try {
    arr.flat(Infinity);
} catch (e) {
    console.log("Flat Circular Recursion Error");
}
"#;
    assert_eq!(run_js(src), vec!["Flat Circular Recursion Error"]);
}

#[test]
fn test_js_array_flatmap_sparse_holes_in_returned_arrays() {
    let src = r#"
const arr = [1];
const res = arr.flatMap(() => [10, , 20]);
console.log(res.length + "|hasHole=" + !(1 in res));
"#;
    assert_eq!(run_js(src), vec!["3|hasHole=true"]);
}

#[test]
fn test_js_array_flat_subclass_species() {
    let src = r#"
class CustomArray extends Array {}
const ca = new CustomArray(1, [2, 3]);
const flat = ca.flat();
console.log(flat.join(",") + "|isCustom=" + (flat instanceof CustomArray));
"#;
    assert_eq!(run_js(src), vec!["1,2,3|isCustom=true"]);
}

#[test]
fn test_js_array_flatmap_subclass_species() {
    let src = r#"
class CustomArray extends Array {}
const ca = new CustomArray(1, 2);
const res = ca.flatMap(x => [x * 2]);
console.log(res.join(",") + "|isCustom=" + (res instanceof CustomArray));
"#;
    assert_eq!(run_js(src), vec!["2,4|isCustom=true"]);
}

#[test]
fn test_js_array_flatmap_skips_input_sparse_holes() {
    let src = r#"
const arr = [1, , 3];
const res = arr.flatMap(x => [x, x * 2]);
console.log(res.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,3,6"]);
}
