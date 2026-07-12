use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// ECMAScript: Typed Arrays — creation, methods, iterators
// ═══════════════════════════════════════════════════════════

// ── Creation ───────────────────────────────────────────────

#[test]
fn int8array_new_with_length() {
    let out = run_js(
        r#"
const a = new Int8Array(3);
console.log(a.length);
console.log(a[0]);
console.log(a[1]);
console.log(a[2]);
"#,
    );
    assert_eq!(out, vec!["3", "0", "0", "0"]);
}

#[test]
fn int8array_from_array_literal() {
    let out = run_js(
        r#"
const a = new Int8Array([10, 20, 30]);
console.log(a[0]);
console.log(a[1]);
console.log(a[2]);
"#,
    );
    assert_eq!(out, vec!["10", "20", "30"]);
}

#[test]
fn uint8clampedarray_clamps_above_255() {
    let out = run_js(
        r#"
const a = new Uint8ClampedArray([0, 128, 255, 300, 1000]);
console.log(a[0]);
console.log(a[2]);
console.log(a[3]);
console.log(a[4]);
"#,
    );
    assert_eq!(out, vec!["0", "255", "255", "255"]);
}

#[test]
fn int8array_overflow_wraps() {
    let out = run_js(
        r#"
const a = new Int8Array([127, 128, 129, -128, -129]);
console.log(a[0]);
console.log(a[1]);
console.log(a[2]);
console.log(a[3]);
console.log(a[4]);
"#,
    );
    assert_eq!(out, vec!["127", "-128", "-127", "-128", "127"]);
}

#[test]
fn float32array_creation_and_element_access() {
    let out = run_js(
        r#"
const a = new Float32Array([1.5, 2.5, 3.5]);
console.log(a[0]);
console.log(a[1]);
console.log(a[2]);
console.log(a.length);
"#,
    );
    assert_eq!(out, vec!["1.5", "2.5", "3.5", "3"]);
}

#[test]
fn float64array_precision() {
    let out = run_js(
        r#"
const a = new Float64Array([3.141592653589793]);
console.log(a[0]);
"#,
    );
    assert_eq!(out, vec!["3.141592653589793"]);
}

#[test]
fn int32array_arithmetic() {
    let out = run_js(
        r#"
const a = new Int32Array([10, 20, 30]);
const sum = a[0] + a[1] + a[2];
console.log(sum);
a[0] = a[1] * 2;
console.log(a[0]);
"#,
    );
    assert_eq!(out, vec!["60", "40"]);
}

#[test]
fn uint32array_length_property() {
    let out = run_js(
        r#"
const a = new Uint32Array(5);
console.log(a.length);
a[4] = 99;
console.log(a[4]);
"#,
    );
    assert_eq!(out, vec!["5", "99"]);
}

// ── Static methods ─────────────────────────────────────────

#[test]
fn typedarray_from_static_method() {
    let out = run_js(
        r#"
const a = Int16Array.from([1, 2, 3, 4]);
console.log(a.length);
console.log(a[0]);
console.log(a[3]);
"#,
    );
    assert_eq!(out, vec!["4", "1", "4"]);
}

#[test]
fn typedarray_of_static_method() {
    let out = run_js(
        r#"
const a = Uint8Array.of(10, 20, 30);
console.log(a.length);
console.log(a[0]);
console.log(a[1]);
console.log(a[2]);
"#,
    );
    assert_eq!(out, vec!["3", "10", "20", "30"]);
}

// ── Instance methods ───────────────────────────────────────

#[test]
fn typedarray_set_method() {
    let out = run_js(
        r#"
const a = new Int32Array(5);
a.set([1, 2, 3], 1);
console.log(a[0]);
console.log(a[1]);
console.log(a[2]);
console.log(a[3]);
"#,
    );
    assert_eq!(out, vec!["0", "1", "2", "3"]);
}

#[test]
fn typedarray_subarray_method() {
    let out = run_js(
        r#"
const a = new Int32Array([1, 2, 3, 4, 5]);
const sub = a.subarray(1, 4);
console.log(sub.length);
console.log(sub[0]);
console.log(sub[2]);
"#,
    );
    assert_eq!(out, vec!["3", "2", "4"]);
}

#[test]
fn typedarray_slice_creates_copy() {
    let out = run_js(
        r#"
const a = new Int32Array([10, 20, 30, 40, 50]);
const copy = a.slice(1, 4);
console.log(copy.length);
console.log(copy[0]);
console.log(copy[2]);
a[1] = 999;
console.log(copy[0]);
"#,
    );
    assert_eq!(out, vec!["3", "20", "40", "20"]);
}

#[test]
fn typedarray_copywithin() {
    let out = run_js(
        r#"
const a = new Int32Array([1, 2, 3, 4, 5]);
a.copyWithin(0, 3);
console.log(a[0]);
console.log(a[1]);
console.log(a[2]);
"#,
    );
    assert_eq!(out, vec!["4", "5", "3"]);
}

#[test]
fn typedarray_fill_method() {
    let out = run_js(
        r#"
const a = new Int32Array(5);
a.fill(7, 1, 4);
console.log(a[0]);
console.log(a[1]);
console.log(a[3]);
console.log(a[4]);
"#,
    );
    assert_eq!(out, vec!["0", "7", "7", "0"]);
}

#[test]
fn typedarray_indexof_method() {
    let out = run_js(
        r#"
const a = new Int32Array([10, 20, 30, 20]);
console.log(a.indexOf(20));
console.log(a.indexOf(99));
"#,
    );
    assert_eq!(out, vec!["1", "-1"]);
}

#[test]
fn typedarray_lastindexof_method() {
    let out = run_js(
        r#"
const a = new Int32Array([10, 20, 30, 20, 10]);
console.log(a.lastIndexOf(20));
console.log(a.lastIndexOf(10));
console.log(a.lastIndexOf(99));
"#,
    );
    assert_eq!(out, vec!["3", "4", "-1"]);
}

#[test]
fn typedarray_includes_method() {
    let out = run_js(
        r#"
const a = new Uint8Array([1, 2, 3, 4, 5]);
console.log(a.includes(3));
console.log(a.includes(6));
"#,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn typedarray_find_method() {
    let out = run_js(
        r#"
const a = new Int32Array([1, 2, 10, 4, 5]);
const val = a.find(x => x > 5);
console.log(val);
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn typedarray_findindex_method() {
    let out = run_js(
        r#"
const a = new Int32Array([1, 2, 10, 4, 5]);
const idx = a.findIndex(x => x > 5);
console.log(idx);
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn typedarray_foreach_method() {
    let out = run_js(
        r#"
const a = new Int32Array([1, 2, 3]);
let sum = 0;
a.forEach(x => { sum += x; });
console.log(sum);
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn typedarray_map_returns_new_typed_array() {
    let out = run_js(
        r#"
const a = new Int32Array([1, 2, 3]);
const b = a.map(x => x * 2);
console.log(b[0]);
console.log(b[1]);
console.log(b[2]);
console.log(b.length);
"#,
    );
    assert_eq!(out, vec!["2", "4", "6", "3"]);
}

#[test]
fn typedarray_filter_returns_new_typed_array() {
    let out = run_js(
        r#"
const a = new Int32Array([1, 2, 3, 4, 5, 6]);
const evens = a.filter(x => x % 2 === 0);
console.log(evens.length);
console.log(evens[0]);
console.log(evens[1]);
console.log(evens[2]);
"#,
    );
    assert_eq!(out, vec!["3", "2", "4", "6"]);
}

#[test]
fn typedarray_reduce_method() {
    let out = run_js(
        r#"
const a = new Int32Array([1, 2, 3, 4, 5]);
const sum = a.reduce((acc, x) => acc + x, 0);
console.log(sum);
"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn typedarray_every_method() {
    let out = run_js(
        r#"
const a = new Int32Array([2, 4, 6, 8]);
console.log(a.every(x => x % 2 === 0));
const b = new Int32Array([2, 3, 6]);
console.log(b.every(x => x % 2 === 0));
"#,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn typedarray_some_method() {
    let out = run_js(
        r#"
const a = new Int32Array([1, 3, 5, 7]);
console.log(a.some(x => x > 4));
console.log(a.some(x => x > 10));
"#,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn typedarray_join_method() {
    let out = run_js(
        r#"
const a = new Int32Array([1, 2, 3, 4]);
console.log(a.join(","));
console.log(a.join("-"));
"#,
    );
    assert_eq!(out, vec!["1,2,3,4", "1-2-3-4"]);
}

#[test]
fn typedarray_reverse_method() {
    let out = run_js(
        r#"
const a = new Int32Array([1, 2, 3, 4, 5]);
a.reverse();
console.log(a[0]);
console.log(a[4]);
console.log(a.join(","));
"#,
    );
    assert_eq!(out, vec!["5", "1", "5,4,3,2,1"]);
}

#[test]
fn typedarray_sort_method() {
    let out = run_js(
        r#"
const a = new Int32Array([3, 1, 4, 1, 5, 9]);
a.sort();
console.log(a.join(","));
"#,
    );
    assert_eq!(out, vec!["1,1,3,4,5,9"]);
}

// ── Iterators ──────────────────────────────────────────────

#[test]
fn typedarray_keys_iterator() {
    let out = run_js(
        r#"
const a = new Int32Array([10, 20, 30]);
const keys = [];
for (const k of a.keys()) {
    keys.push(k);
}
console.log(keys.join(","));
"#,
    );
    assert_eq!(out, vec!["0,1,2"]);
}

#[test]
fn typedarray_values_iterator() {
    let out = run_js(
        r#"
const a = new Int32Array([10, 20, 30]);
const vals = [];
for (const v of a.values()) {
    vals.push(v);
}
console.log(vals.join(","));
"#,
    );
    assert_eq!(out, vec!["10,20,30"]);
}

#[test]
fn typedarray_entries_iterator() {
    let out = run_js(
        r#"
const a = new Int32Array([10, 20, 30]);
const pairs = [];
for (const [i, v] of a.entries()) {
    pairs.push(i + ":" + v);
}
console.log(pairs.join(","));
"#,
    );
    assert_eq!(out, vec!["0:10,1:20,2:30"]);
}

// ── Properties ─────────────────────────────────────────────

#[test]
fn typedarray_bytes_per_element_property() {
    let out = run_js(
        r#"
console.log(Int8Array.BYTES_PER_ELEMENT);
console.log(Int16Array.BYTES_PER_ELEMENT);
console.log(Int32Array.BYTES_PER_ELEMENT);
console.log(Float64Array.BYTES_PER_ELEMENT);
"#,
    );
    assert_eq!(out, vec!["1", "2", "4", "8"]);
}

// ── ArrayBuffer sharing ────────────────────────────────────

#[test]
fn shared_arraybuffer_between_two_typed_arrays() {
    let out = run_js(
        r#"
const buf = new ArrayBuffer(8);
const i32 = new Int32Array(buf);
const u8  = new Uint8Array(buf);
i32[0] = 1;
console.log(u8[0]);
i32[0] = 256;
console.log(u8[0]);
console.log(u8[1]);
"#,
    );
    assert_eq!(out, vec!["1", "0", "1"]);
}

// ── BigInt typed arrays ────────────────────────────────────

#[test]
fn bigint64array_creation_with_bigint_values() {
    let out = run_js(
        r#"
const a = new BigInt64Array([1n, 2n, 9007199254740993n]);
console.log(a[0]);
console.log(a[1]);
console.log(a[2]);
console.log(a.length);
"#,
    );
    // Node-verified: BigInt64Array elements print with the `n` suffix;
    // length is a plain number.
    assert_eq!(out, vec!["1n", "2n", "9007199254740993n", "3"]);
}
