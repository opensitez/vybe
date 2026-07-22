use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: DataView Binary Read/Write & Endianness Mechanics
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_dataview_get_set_int8_uint8() {
    let src = r#"
const buffer = new ArrayBuffer(4);
const dv = new DataView(buffer);
dv.setInt8(0, -128);
dv.setUint8(1, 255);

console.log(`${dv.getInt8(0)}:${dv.getUint8(1)}`);
"#;
    assert_eq!(run_js(src), vec!["-128:255"]);
}

#[test]
fn test_js_dataview_get_set_int16_little_vs_big_endian() {
    let src = r#"
const buffer = new ArrayBuffer(2);
const dv = new DataView(buffer);
dv.setInt16(0, 0x1234, false); // Big-Endian write: 0x12, 0x34

console.log(`BigEndian=0x${dv.getInt16(0, false).toString(16)}|LittleEndian=0x${dv.getInt16(0, true).toString(16)}`);
"#;
    assert_eq!(run_js(src), vec!["BigEndian=0x1234|LittleEndian=0x3412"]);
}

#[test]
fn test_js_dataview_get_set_uint32_little_endian() {
    let src = r#"
const buffer = new ArrayBuffer(4);
const dv = new DataView(buffer);
dv.setUint32(0, 0xDEADBEEF, true); // Little endian write

console.log(dv.getUint32(0, true).toString(16).toUpperCase());
"#;
    assert_eq!(run_js(src), vec!["DEADBEEF"]);
}

#[test]
fn test_js_dataview_get_set_float32_float64() {
    let src = r#"
const buffer = new ArrayBuffer(12);
const dv = new DataView(buffer);
dv.setFloat32(0, 3.14, true);
dv.setFloat64(4, 2.718281828459045, false);

console.log(dv.getFloat32(0, true).toFixed(2) + "|" + dv.getFloat64(4, false));
"#;
    assert_eq!(run_js(src), vec!["3.14|2.718281828459045"]);
}

#[test]
fn test_js_dataview_get_set_bigint64_biguint64() {
    let src = r#"
const buffer = new ArrayBuffer(16);
const dv = new DataView(buffer);
dv.setBigInt64(0, -9007199254740991n, true);
dv.setBigUint64(8, 18446744073709551615n, false);

console.log(dv.getBigInt64(0, true).toString() + "|" + dv.getBigUint64(8, false).toString());
"#;
    assert_eq!(run_js(src), vec!["-9007199254740991|18446744073709551615"]);
}

#[test]
fn test_js_dataview_out_of_bounds_read_throws_rangeerror() {
    let src = r#"
const buffer = new ArrayBuffer(4);
const dv = new DataView(buffer);
try {
    dv.getInt32(2); // Reads 4 bytes starting at offset 2 (exceeds buffer length 4)!
} catch (e) {
    console.log("DataView Read OutOfBounds RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["DataView Read OutOfBounds RangeError"]);
}

#[test]
fn test_js_dataview_out_of_bounds_write_throws_rangeerror() {
    let src = r#"
const buffer = new ArrayBuffer(4);
const dv = new DataView(buffer);
try {
    dv.setUint32(1, 100);
} catch (e) {
    console.log("DataView Write OutOfBounds RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["DataView Write OutOfBounds RangeError"]);
}

#[test]
fn test_js_dataview_byte_length_and_byte_offset_properties() {
    let src = r#"
const buffer = new ArrayBuffer(16);
const dv = new DataView(buffer, 4, 8);
console.log(`${dv.byteOffset}:${dv.byteLength}:${dv.buffer.byteLength}`);
"#;
    assert_eq!(run_js(src), vec!["4:8:16"]);
}

#[test]
fn test_js_dataview_default_endianness_is_big_endian() {
    let src = r#"
const buffer = new ArrayBuffer(2);
const dv = new DataView(buffer);
dv.setUint16(0, 0x0102); // Omitted littleEndian flag defaults to FALSE (Big-Endian)
const u8 = new Uint8Array(buffer);
console.log(`${u8[0]}:${u8[1]}`);
"#;
    assert_eq!(run_js(src), vec!["1:2"]);
}

#[test]
fn test_js_dataview_get_set_float16_es2024() {
    let src = r#"
const buffer = new ArrayBuffer(2);
const dv = new DataView(buffer);
dv.setFloat16(0, 1.5, true);
console.log(dv.getFloat16(0, true));
"#;
    assert_eq!(run_js(src), vec!["1.5"]);
}

#[test]
fn test_js_dataview_detached_buffer_access_throws_typeerror() {
    let src = r#"
const buffer = new ArrayBuffer(8);
const dv = new DataView(buffer);
buffer.transfer();
try {
    dv.getUint8(0);
} catch (e) {
    console.log("DataView Detached Buffer TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["DataView Detached Buffer TypeError"]);
}

#[test]
fn test_js_dataview_unaligned_offsets_supported() {
    let src = r#"
const buffer = new ArrayBuffer(8);
const dv = new DataView(buffer);
dv.setUint32(1, 0x12345678, true); // DataView supports unaligned byte offsets!
console.log(dv.getUint32(1, true).toString(16));
"#;
    assert_eq!(run_js(src), vec!["12345678"]);
}

#[test]
fn test_js_dataview_constructor_offset_exceeds_buffer_length_throws_rangeerror() {
    let src = r#"
const buffer = new ArrayBuffer(8);
try {
    new DataView(buffer, 16);
} catch (e) {
    console.log("DataView Constructor Offset RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["DataView Constructor Offset RangeError"]);
}

#[test]
fn test_js_dataview_constructor_non_arraybuffer_throws_typeerror() {
    let src = r#"
try {
    new DataView({});
} catch (e) {
    console.log("DataView Non-ArrayBuffer TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["DataView Non-ArrayBuffer TypeError"]);
}

#[test]
fn test_js_dataview_negative_offset_throws_rangeerror() {
    let src = r#"
const buffer = new ArrayBuffer(8);
const dv = new DataView(buffer);
try {
    dv.getUint8(-1);
} catch (e) {
    console.log("Negative Offset RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["Negative Offset RangeError"]);
}

#[test]
fn test_js_dataview_bigint_conversion_typeerror() {
    let src = r#"
const buffer = new ArrayBuffer(8);
const dv = new DataView(buffer);
try {
    dv.setBigInt64(0, 12345); // Passing regular Number to BigInt method throws TypeError!
} catch (e) {
    console.log("BigInt DataView Conversion TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["BigInt DataView Conversion TypeError"]);
}

#[test]
fn test_js_dataview_number_conversion_typeerror() {
    let src = r#"
const buffer = new ArrayBuffer(8);
const dv = new DataView(buffer);
try {
    dv.setUint32(0, 100n); // Passing BigInt to Number method throws TypeError!
} catch (e) {
    console.log("Number DataView Conversion TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Number DataView Conversion TypeError"]);
}

#[test]
fn test_js_dataview_resizable_buffer_auto_byte_length() {
    let src = r#"
const buffer = new ArrayBuffer(8, { maxByteLength: 32 });
const dv = new DataView(buffer);
console.log(dv.byteLength);
buffer.resize(16);
console.log(dv.byteLength);
"#;
    assert_eq!(run_js(src), vec!["8", "16"]);
}

#[test]
fn test_js_dataview_nan_floating_point_write_read() {
    let src = r#"
const buffer = new ArrayBuffer(4);
const dv = new DataView(buffer);
dv.setFloat32(0, NaN);
console.log(Number.isNaN(dv.getFloat32(0)));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_dataview_prototype_symbol_tostringtag() {
    let src = r#"
const buffer = new ArrayBuffer(8);
const dv = new DataView(buffer);
console.log(Object.prototype.toString.call(dv));
"#;
    assert_eq!(run_js(src), vec!["[object DataView]"]);
}
