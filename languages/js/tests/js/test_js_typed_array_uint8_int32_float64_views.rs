use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: TypedArray Constructors & Element Access (Uint8Array, Int32Array, Float64Array)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_typedarray_uint8_overflow_wrap() {
    let src = r#"
const u8 = new Uint8Array(2);
u8[0] = 255;
u8[1] = 256; // Wraps modulo 256
console.log(`${u8[0]}:${u8[1]}`);
"#;
    assert_eq!(run_js(src), vec!["255:0"]);
}

#[test]
fn test_js_typedarray_int32_signed_overflow() {
    let src = r#"
const i32 = new Int32Array(1);
i32[0] = 2147483647 + 1; // Signed 32-bit wrap
console.log(i32[0]);
"#;
    assert_eq!(run_js(src), vec!["-2147483648"]);
}

#[test]
fn test_js_typedarray_float64_precision() {
    let src = r#"
const f64 = new Float64Array(2);
f64[0] = 3.141592653589793;
f64[1] = NaN;
console.log(f64[0] + "|" + Number.isNaN(f64[1]));
"#;
    assert_eq!(run_js(src), vec!["3.141592653589793|true"]);
}

#[test]
fn test_js_typedarray_bytes_per_element_constants() {
    let src = r#"
console.log(`${Uint8Array.BYTES_PER_ELEMENT}:${Int32Array.BYTES_PER_ELEMENT}:${Float64Array.BYTES_PER_ELEMENT}`);
"#;
    assert_eq!(run_js(src), vec!["1:4:8"]);
}

#[test]
fn test_js_typedarray_buffer_and_byteoffset_byte_length() {
    let src = r#"
const buffer = new ArrayBuffer(16);
const view = new Int32Array(buffer, 4, 2);
console.log(`${view.length}:${view.byteOffset}:${view.byteLength}`);
"#;
    assert_eq!(run_js(src), vec!["2:4:8"]);
}

#[test]
fn test_js_typedarray_uint8clamped_clamping_behavior() {
    let src = r#"
const clamped = new Uint8ClampedArray(3);
clamped[0] = -10; // Clamped to 0
clamped[1] = 300; // Clamped to 255
clamped[2] = 2.5; // Rounding half to even: 2
console.log(clamped.join(","));
"#;
    assert_eq!(run_js(src), vec!["0,255,2"]);
}

#[test]
fn test_js_typedarray_view_underlying_buffer_shared_memory() {
    let src = r#"
const buf = new ArrayBuffer(4);
const u8 = new Uint8Array(buf);
const u32 = new Uint32Array(buf);

u8[0] = 0xFF;
console.log(u32[0] !== 0); // Modifying u8 updates shared u32 view in buffer
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_typedarray_out_of_bounds_index_returns_undefined() {
    let src = r#"
const arr = new Uint8Array(2);
console.log(arr[2] === undefined + "|" + arr[-1] === undefined);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_typedarray_out_of_bounds_assignment_ignored() {
    let src = r#"
const arr = new Uint8Array(2);
arr[5] = 100;
console.log(arr.length + "|" + (arr[5] === undefined));
"#;
    assert_eq!(run_js(src), vec!["2|true"]);
}

#[test]
fn test_js_typedarray_iterator_protocol() {
    let src = r#"
const u8 = new Uint8Array([10, 20, 30]);
console.log([...u8].join("-"));
"#;
    assert_eq!(run_js(src), vec!["10-20-30"]);
}

#[test]
fn test_js_typedarray_for_in_loop_only_enumerates_numeric_indices() {
    let src = r#"
const arr = new Uint8Array([5, 10]);
const keys = [];
for (const k in arr) keys.push(k);
console.log(keys.join(","));
"#;
    assert_eq!(run_js(src), vec!["0,1"]);
}

#[test]
fn test_js_typedarray_constructor_with_array_like() {
    let src = r#"
const arr = new Int16Array([100, 200, 300]);
console.log(arr.length + "|" + arr[1]);
"#;
    assert_eq!(run_js(src), vec!["3|200"]);
}

#[test]
fn test_js_typedarray_constructor_copy_from_another_typedarray() {
    let src = r#"
const original = new Float32Array([1.5, 2.5]);
const copy = new Float64Array(original);
console.log(copy.length + "|" + copy[0] + "|" + (copy.buffer !== original.buffer));
"#;
    assert_eq!(run_js(src), vec!["2|1.5|true"]);
}

#[test]
fn test_js_typedarray_unaligned_byte_offset_throws_rangeerror() {
    let src = r#"
const buf = new ArrayBuffer(16);
try {
    new Int32Array(buf, 3); // ByteOffset 3 is not a multiple of Int32 element size (4)!
} catch (e) {
    console.log("Unaligned ByteOffset RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["Unaligned ByteOffset RangeError"]);
}

#[test]
fn test_js_typedarray_bigint64_view_support() {
    let src = r#"
const big64 = new BigInt64Array(2);
big64[0] = 9007199254740991n;
console.log(big64[0].toString());
"#;
    assert_eq!(run_js(src), vec!["9007199254740991"]);
}

#[test]
fn test_js_typedarray_biguint64_view_overflow() {
    let src = r#"
const bigu64 = new BigUint64Array(1);
bigu64[0] = 0xFFFFFFFFFFFFFFFFn;
console.log(bigu64[0].toString());
"#;
    assert_eq!(run_js(src), vec!["18446744073709551615"]);
}

#[test]
fn test_js_typedarray_fill_method() {
    let src = r#"
const arr = new Uint8Array(4);
arr.fill(42, 1, 3);
console.log(arr.join(","));
"#;
    assert_eq!(run_js(src), vec!["0,42,42,0"]);
}

#[test]
fn test_js_typedarray_map_returns_same_typedarray_constructor() {
    let src = r#"
const u8 = new Uint8Array([1, 2, 3]);
const mapped = u8.map(x => x * 10);
console.log(mapped.join(",") + "|isUint8=" + (mapped instanceof Uint8Array));
"#;
    assert_eq!(run_js(src), vec!["10,20,30|isUint8=true"]);
}

#[test]
fn test_js_typedarray_cannot_delete_indexed_properties() {
    let src = r#"
const arr = new Uint8Array([10]);
try {
    "use strict";
    delete arr[0];
} catch (e) {
    console.log("Delete TypedArray Index TypeError");
}
console.log(arr[0]);
"#;
    assert_eq!(run_js(src), vec!["Delete TypedArray Index TypeError", "10"]);
}

#[test]
fn test_js_typedarray_symbol_species_override() {
    let src = r#"
class CustomUint8 extends Uint8Array {}
const cu8 = new CustomUint8([5, 10]);
const sliced = cu8.slice(0, 1);
console.log(sliced[0] + "|isCustom=" + (sliced instanceof CustomUint8));
"#;
    assert_eq!(run_js(src), vec!["5|isCustom=true"]);
}
