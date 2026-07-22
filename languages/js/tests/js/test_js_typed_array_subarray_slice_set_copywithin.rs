use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: TypedArray Subarray, Slice, Set & CopyWithin Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_typedarray_subarray_shares_underlying_buffer() {
    let src = r#"
const original = new Uint8Array([10, 20, 30, 40]);
const sub = original.subarray(1, 3);
sub[0] = 99; // Modifying subarray updates original array!

console.log(original.join(",") + "|subLen=" + sub.length);
"#;
    assert_eq!(run_js(src), vec!["10,99,30,40|subLen=2"]);
}

#[test]
fn test_js_typedarray_slice_copies_underlying_buffer() {
    let src = r#"
const original = new Uint8Array([10, 20, 30, 40]);
const sliced = original.slice(1, 3);
sliced[0] = 99; // Modifying sliced copy does NOT update original!

console.log(original.join(",") + "|" + sliced.join(","));
"#;
    assert_eq!(run_js(src), vec!["10,20,30,40|99,30"]);
}

#[test]
fn test_js_typedarray_set_copy_from_array() {
    let src = r#"
const dest = new Uint8Array(5);
dest.set([10, 20, 30], 1); // Set starting at index 1
console.log(dest.join(","));
"#;
    assert_eq!(run_js(src), vec!["0,10,20,30,0"]);
}

#[test]
fn test_js_typedarray_set_copy_from_typedarray() {
    let src = r#"
const srcArr = new Uint8Array([5, 15]);
const dest = new Uint8Array(4);
dest.set(srcArr, 2);
console.log(dest.join(","));
"#;
    assert_eq!(run_js(src), vec!["0,0,5,15"]);
}

#[test]
fn test_js_typedarray_set_out_of_bounds_throws_rangeerror() {
    let src = r#"
const dest = new Uint8Array(3);
try {
    dest.set([1, 2, 3], 2); // Exceeds length 3!
} catch (e) {
    console.log("TypedArray Set RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["TypedArray Set RangeError"]);
}

#[test]
fn test_js_typedarray_copywithin_internal_buffer_copy() {
    let src = r#"
const arr = new Uint8Array([1, 2, 3, 4, 5]);
arr.copyWithin(0, 3, 5); // Copy elements at 3..5 to index 0
console.log(arr.join(","));
"#;
    assert_eq!(run_js(src), vec!["4,5,3,4,5"]);
}

#[test]
fn test_js_typedarray_copywithin_overlapping_region_safety() {
    let src = r#"
const arr = new Uint8Array([1, 2, 3, 4, 5]);
arr.copyWithin(1, 0, 3); // Copy 0..3 to index 1 (overlapping)
console.log(arr.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,1,2,3,5"]);
}

#[test]
fn test_js_typedarray_subarray_negative_indices() {
    let src = r#"
const arr = new Int32Array([10, 20, 30, 40]);
const sub = arr.subarray(-3, -1);
console.log(sub.join(","));
"#;
    assert_eq!(run_js(src), vec!["20,30"]);
}

#[test]
fn test_js_typedarray_slice_negative_indices() {
    let src = r#"
const arr = new Int32Array([10, 20, 30, 40]);
const sliced = arr.slice(-2);
console.log(sliced.join(","));
"#;
    assert_eq!(run_js(src), vec!["30,40"]);
}

#[test]
fn test_js_typedarray_set_negative_offset_throws_rangeerror() {
    let src = r#"
const arr = new Uint8Array(4);
try {
    arr.set([1], -1);
} catch (e) {
    console.log("Negative Offset RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["Negative Offset RangeError"]);
}

#[test]
fn test_js_typedarray_subarray_omitted_end_slices_to_end() {
    let src = r#"
const arr = new Uint8Array([1, 2, 3, 4]);
console.log(arr.subarray(2).join(","));
"#;
    assert_eq!(run_js(src), vec!["3,4"]);
}

#[test]
fn test_js_typedarray_slice_omitted_end_slices_to_end() {
    let src = r#"
const arr = new Uint8Array([1, 2, 3, 4]);
console.log(arr.slice(2).join(","));
"#;
    assert_eq!(run_js(src), vec!["3,4"]);
}

#[test]
fn test_js_typedarray_set_coercion_of_input_elements() {
    let src = r#"
const arr = new Uint8Array(2);
arr.set(["10", 256]); // "10" -> 10, 256 -> 0
console.log(arr.join(","));
"#;
    assert_eq!(run_js(src), vec!["10,0"]);
}

#[test]
fn test_js_typedarray_copywithin_negative_target() {
    let src = r#"
const arr = new Uint8Array([10, 20, 30, 40]);
arr.copyWithin(-2, 0, 2);
console.log(arr.join(","));
"#;
    assert_eq!(run_js(src), vec!["10,20,10,20"]);
}

#[test]
fn test_js_typedarray_subarray_returns_same_typedarray_constructor() {
    let src = r#"
const i32 = new Int32Array([1, 2, 3]);
const sub = i32.subarray(1);
console.log(sub instanceof Int32Array);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_typedarray_slice_returns_new_buffer() {
    let src = r#"
const u8 = new Uint8Array([1, 2]);
const sliced = u8.slice();
console.log(sliced.buffer !== u8.buffer);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_typedarray_set_overlapping_same_buffer() {
    let src = r#"
const arr = new Uint8Array([1, 2, 3, 4]);
arr.set(arr.subarray(0, 2), 2); // Copy [1,2] to index 2
console.log(arr.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,1,2"]);
}

#[test]
fn test_js_typedarray_copywithin_start_greater_than_end_no_op() {
    let src = r#"
const arr = new Uint8Array([1, 2, 3]);
arr.copyWithin(0, 2, 1);
console.log(arr.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,3"]);
}

#[test]
fn test_js_typedarray_subarray_zero_length() {
    let src = r#"
const arr = new Uint8Array([1, 2, 3]);
const sub = arr.subarray(2, 2);
console.log(sub.length);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_typedarray_set_bigint_view_type_safety() {
    let src = r#"
const big = new BigInt64Array(2);
big.set([100n, 200n]);
console.log(big[0].toString() + "|" + big[1].toString());
"#;
    assert_eq!(run_js(src), vec!["100|200"]);
}
