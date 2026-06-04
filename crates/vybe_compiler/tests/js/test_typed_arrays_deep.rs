/// Typed array operations — binary data, views, conversions
use super::helpers::run_js;

#[test]
fn typed_array_creation_methods() {
    assert_eq!(
        run_js(
            r#"
const a = new Int32Array([1, 2, 3, 4, 5]);
const b = Int32Array.of(10, 20, 30);
const c = Int32Array.from([1.5, 2.7, 3.9]);
console.log(a.length);
console.log(b[1]);
console.log(c[0]);
"#
        ),
        vec!["5", "20", "1"]
    );
}

#[test]
fn typed_array_shared_buffer() {
    assert_eq!(
        run_js(
            r#"
const buffer = new ArrayBuffer(16);
const int32 = new Int32Array(buffer);
int32[0] = 1;
int32[1] = 256;
console.log(int32[0]);
console.log(int32[1]);
console.log(buffer.byteLength);
"#
        ),
        vec!["1", "256", "16"]
    );
}

#[test]
fn typed_array_slice_copy() {
    assert_eq!(
        run_js(
            r#"
const orig = new Float32Array([1.0, 2.0, 3.0, 4.0]);
const sliced = orig.slice(1, 3);
sliced[0] = 99;
console.log(orig[1]);
console.log(sliced[0]);
console.log(sliced.length);
"#
        ),
        vec!["2", "99", "2"]
    );
}

#[test]
fn typed_array_set_method() {
    assert_eq!(
        run_js(
            r#"
const dest = new Int32Array(8);
const src = [1, 2, 3];
dest.set(src, 2);
console.log(dest[0]);
console.log(dest[2]);
console.log(dest[3]);
"#
        ),
        vec!["0", "1", "2"]
    );
}

#[test]
fn typed_array_filter_map() {
    assert_eq!(
        run_js(
            r#"
const arr = new Int32Array([1, 2, 3, 4, 5, 6]);
const evens = arr.filter(x => x % 2 === 0);
const doubled = arr.map(x => x * 2);
console.log(evens.join(","));
console.log(doubled.join(","));
"#
        ),
        vec!["2,4,6", "2,4,6,8,10,12"]
    );
}

#[test]
fn dataview_read_write() {
    assert_eq!(
        run_js(
            r#"
const buf = new ArrayBuffer(8);
const view = new DataView(buf);
view.setFloat64(0, Math.PI, false);  // big endian
const pi = view.getFloat64(0, false);
console.log(Math.abs(pi - Math.PI) < 1e-10);
view.setInt16(0, -1000, true);  // little endian
console.log(view.getInt16(0, true));
"#
        ),
        vec!["true", "-1000"]
    );
}

#[test]
fn uint8clampedarray_clamps() {
    assert_eq!(
        run_js(
            r#"
const arr = new Uint8ClampedArray(4);
arr[0] = 300;
arr[1] = -10;
arr[2] = 128;
arr[3] = 0;
console.log(arr[0]);
console.log(arr[1]);
console.log(arr[2]);
console.log(arr[3]);
"#
        ),
        vec!["255", "0", "128", "0"]
    );
}

#[test]
fn typed_array_reduce() {
    assert_eq!(
        run_js(
            r#"
const arr = new Float64Array([1.5, 2.5, 3.0, 4.0]);
const sum = arr.reduce((a, b) => a + b, 0);
const max = arr.reduce((a, b) => Math.max(a, b), -Infinity);
console.log(sum);
console.log(max);
"#
        ),
        vec!["11", "4"]
    );
}

#[test]
fn arraybuffer_transfer_copy() {
    assert_eq!(
        run_js(
            r#"
const buf1 = new ArrayBuffer(4);
const view1 = new Uint32Array(buf1);
view1[0] = 42;
// Copy via typed array
const buf2 = buf1.slice(0);
const view2 = new Uint32Array(buf2);
view2[0] = 99;
console.log(view1[0]);
console.log(view2[0]);
"#
        ),
        vec!["42", "99"]
    );
}

#[test]
fn typed_array_sort_find() {
    assert_eq!(
        run_js(
            r#"
const arr = new Int32Array([5, 3, 1, 4, 2]);
arr.sort();
console.log(Array.from(arr).join(","));
console.log(arr.find(x => x > 3));
console.log(arr.findIndex(x => x > 3));
"#
        ),
        vec!["1,2,3,4,5", "4", "3"]
    );
}

#[test]
fn float32_precision_loss() {
    assert_eq!(
        run_js(
            r#"
const f64 = 1.337;
const arr = new Float32Array(1);
arr[0] = f64;
const f32 = arr[0];
console.log(f32 !== f64);
console.log(Math.abs(f32 - f64) < 0.0001);
"#
        ),
        vec!["true", "true"]
    );
}
