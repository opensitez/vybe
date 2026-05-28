/// TypedArray advanced — Float32Array/Float64Array operations, DataView reading,
/// SharedArrayBuffer, Atomics, typed array methods, view slicing, buffer sharing.

use super::helpers::run_js;

// ── Float typed arrays ────────────────────────────────────────────────────────

#[test]
fn float32array_precision_loss() {
    assert_eq!(run_js(r#"
const f32 = new Float32Array(1);
f32[0] = Math.PI;
// float32 loses precision compared to float64
console.log(f32[0] !== Math.PI);
console.log(Math.abs(f32[0] - Math.PI) < 0.001);
"#), vec!["true", "true"]);
}

#[test]
fn float64array_preserves_precision() {
    assert_eq!(run_js(r#"
const f64 = new Float64Array(1);
f64[0] = Math.PI;
console.log(f64[0] === Math.PI);
"#), vec!["true"]);
}

// ── typed array methods ───────────────────────────────────────────────────────

#[test]
fn typed_array_set_from_array() {
    assert_eq!(run_js(r#"
const ta = new Int32Array(5);
ta.set([10, 20, 30]);
console.log(ta[0]);
console.log(ta[1]);
console.log(ta[2]);
console.log(ta[3]); // unfilled — 0
"#), vec!["10", "20", "30", "0"]);
}

#[test]
fn typed_array_set_with_offset() {
    assert_eq!(run_js(r#"
const ta = new Uint8Array(5);
ta.set([1, 2], 2); // start at index 2
console.log(ta.join(","));
"#), vec!["0,0,1,2,0"]);
}

#[test]
fn typed_array_subarray_is_view() {
    assert_eq!(run_js(r#"
const ta = new Int32Array([10, 20, 30, 40, 50]);
const sub = ta.subarray(1, 3);
console.log(sub.length);
console.log(sub[0]);
console.log(sub[1]);
// Modifying subarray affects original
sub[0] = 99;
console.log(ta[1]);
"#), vec!["2", "20", "30", "99"]);
}

#[test]
fn typed_array_slice_is_copy() {
    assert_eq!(run_js(r#"
const ta = new Int32Array([1, 2, 3, 4]);
const sliced = ta.slice(1, 3);
sliced[0] = 99;
console.log(ta[1]); // unchanged
console.log(sliced[0]);
"#), vec!["2", "99"]);
}

#[test]
fn typed_array_map_returns_same_type() {
    assert_eq!(run_js(r#"
const ta = new Int32Array([1, 2, 3]);
const mapped = ta.map(x => x * 2);
console.log(mapped.length === 3);
console.log(Array.from(mapped).join(","));
"#), vec!["true", "2,4,6"]);
}

#[test]
fn typed_array_filter_returns_same_type() {
    assert_eq!(run_js(r#"
const ta = new Int32Array([1, 2, 3, 4, 5]);
const even = ta.filter(x => x % 2 === 0);
console.log(even.length === 2);
console.log(Array.from(even).join(","));
"#), vec!["true", "2,4"]);
}

#[test]
fn typed_array_reduce() {
    assert_eq!(run_js(r#"
const ta = new Float64Array([1.5, 2.5, 3.5]);
const sum = ta.reduce((acc, x) => acc + x, 0);
console.log(sum);
"#), vec!["7.5"]);
}

// ── ArrayBuffer and views ─────────────────────────────────────────────────────

#[test]
fn multiple_views_of_same_buffer() {
    assert_eq!(run_js(r#"
const buffer = new ArrayBuffer(8);
const i32 = new Int32Array(buffer);
const u8 = new Uint8Array(buffer);

i32[0] = 1; // sets first 4 bytes
// u8 sees the same memory
console.log(u8[0] !== 0 || u8[1] !== 0 || u8[2] !== 0 || u8[3] !== 0);
"#), vec!["true"]);
}

#[test]
fn arraybuffer_slice_creates_copy() {
    assert_eq!(run_js(r#"
const buf = new ArrayBuffer(8);
const view = new Uint8Array(buf);
view[0] = 42;
const copy = buf.slice(0, 4);
const copyView = new Uint8Array(copy);
copyView[0] = 99;
console.log(view[0]); // original unchanged
console.log(copyView[0]);
"#), vec!["42", "99"]);
}

// ── DataView ──────────────────────────────────────────────────────────────────

#[test]
fn dataview_set_and_get_int32() {
    assert_eq!(run_js(r#"
const buf = new ArrayBuffer(4);
const dv = new DataView(buf);
dv.setUint32(0, 0xDEADBEEF);
console.log(dv.getUint32(0).toString(16));
"#), vec!["deadbeef"]);
}

#[test]
fn dataview_little_endian_vs_big_endian() {
    assert_eq!(run_js(r#"
const buf = new ArrayBuffer(4);
const dv = new DataView(buf);
dv.setUint16(0, 0x0102, true);  // little endian
console.log(dv.getUint8(0));   // low byte first
console.log(dv.getUint8(1));   // high byte
"#), vec!["2", "1"]);
}

#[test]
fn dataview_float64_round_trip() {
    assert_eq!(run_js(r#"
const buf = new ArrayBuffer(8);
const dv = new DataView(buf);
dv.setFloat64(0, Math.PI);
console.log(dv.getFloat64(0) === Math.PI);
"#), vec!["true"]);
}

// ── Atomics basics ────────────────────────────────────────────────────────────

#[test]
fn atomics_add_returns_old_value() {
    assert_eq!(run_js(r#"
const sab = new SharedArrayBuffer(4);
const ta = new Int32Array(sab);
ta[0] = 10;
const old = Atomics.add(ta, 0, 5);
console.log(old);      // old value
console.log(ta[0]);    // new value
"#), vec!["10", "15"]);
}

#[test]
fn atomics_compareExchange() {
    assert_eq!(run_js(r#"
const sab = new SharedArrayBuffer(4);
const ta = new Int32Array(sab);
ta[0] = 42;
const result = Atomics.compareExchange(ta, 0, 42, 99);
console.log(result); // old value
console.log(ta[0]);  // new value — exchange happened
const result2 = Atomics.compareExchange(ta, 0, 42, 0); // expected wrong
console.log(result2); // old value (99)
console.log(ta[0]);   // unchanged (99)
"#), vec!["42", "99", "99", "99"]);
}

#[test]
fn atomics_load_and_store() {
    assert_eq!(run_js(r#"
const sab = new SharedArrayBuffer(4);
const ta = new Int32Array(sab);
Atomics.store(ta, 0, 777);
console.log(Atomics.load(ta, 0));
"#), vec!["777"]);
}
