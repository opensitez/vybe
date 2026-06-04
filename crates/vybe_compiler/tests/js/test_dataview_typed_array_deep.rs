/// DataView, ArrayBuffer, typed array interactions
use super::helpers::run_js;

#[test]
fn dataview_read_write_int8() {
    assert_eq!(
        run_js(
            r#"
const buf = new ArrayBuffer(4);
const view = new DataView(buf);
view.setInt8(0, -1);
view.setInt8(1, 127);
console.log(view.getInt8(0));
console.log(view.getInt8(1));
"#
        ),
        vec!["-1", "127"]
    );
}

#[test]
fn dataview_read_write_uint16_endianness() {
    assert_eq!(
        run_js(
            r#"
const buf = new ArrayBuffer(2);
const view = new DataView(buf);
view.setUint16(0, 0x0102, true); // little-endian
console.log(view.getUint8(0)); // low byte
console.log(view.getUint8(1)); // high byte
"#
        ),
        vec!["2", "1"]
    );
}

#[test]
fn dataview_float64_roundtrip() {
    assert_eq!(
        run_js(
            r#"
const buf = new ArrayBuffer(8);
const view = new DataView(buf);
const pi = Math.PI;
view.setFloat64(0, pi);
console.log(view.getFloat64(0) === pi);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn typed_array_shared_buffer() {
    assert_eq!(
        run_js(
            r#"
const buf = new ArrayBuffer(8);
const i32 = new Int32Array(buf);
const u8 = new Uint8Array(buf);
i32[0] = 0x01020304;
// u8 sees the bytes of i32[0]
const bytes = [u8[0], u8[1], u8[2], u8[3]];
console.log(bytes.some(b => b !== 0));
"#
        ),
        vec!["true"]
    );
}

#[test]
fn int32array_operations() {
    assert_eq!(
        run_js(
            r#"
const arr = new Int32Array([1, 2, 3, 4, 5]);
const sum = arr.reduce((acc, x) => acc + x, 0);
console.log(sum);
console.log(arr.length);
"#
        ),
        vec!["15", "5"]
    );
}

#[test]
fn float32_precision_loss() {
    assert_eq!(
        run_js(
            r#"
const f32 = new Float32Array(1);
f32[0] = 1.337;
// Float32 has less precision than Float64
console.log(f32[0] !== 1.337);
// But it's close
console.log(Math.abs(f32[0] - 1.337) < 0.001);
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn arraybuffer_slice_is_copy() {
    assert_eq!(
        run_js(
            r#"
const buf = new ArrayBuffer(8);
const view = new Uint8Array(buf);
view[0] = 42;
const slice = buf.slice(0, 4);
const sliceView = new Uint8Array(slice);
sliceView[0] = 99;
console.log(view[0]); // original unchanged
console.log(sliceView[0]);
"#
        ),
        vec!["42", "99"]
    );
}

#[test]
fn typed_array_from_array() {
    assert_eq!(
        run_js(
            r#"
const arr = Int32Array.from([1, 2, 3, 4]);
console.log(arr[0]);
console.log(arr.length);
console.log(arr instanceof Int32Array);
"#
        ),
        vec!["1", "4", "true"]
    );
}

#[test]
fn typed_array_of() {
    assert_eq!(
        run_js(
            r#"
const arr = Float64Array.of(1.1, 2.2, 3.3);
console.log(arr.length);
console.log(arr[0]);
"#
        ),
        vec!["3", "1.1"]
    );
}

#[test]
fn dataview_offset_and_length() {
    assert_eq!(
        run_js(
            r#"
const buf = new ArrayBuffer(10);
const view = new DataView(buf, 2, 4); // offset 2, length 4
console.log(view.byteOffset);
console.log(view.byteLength);
view.setInt8(0, 42); // relative to offset
console.log(view.getInt8(0));
"#
        ),
        vec!["2", "4", "42"]
    );
}

#[test]
fn uint8clampedarray_clamps_values() {
    assert_eq!(
        run_js(
            r#"
const arr = new Uint8ClampedArray(3);
arr[0] = 300;  // clamped to 255
arr[1] = -10;  // clamped to 0
arr[2] = 128;  // unchanged
console.log(arr[0]);
console.log(arr[1]);
console.log(arr[2]);
"#
        ),
        vec!["255", "0", "128"]
    );
}
