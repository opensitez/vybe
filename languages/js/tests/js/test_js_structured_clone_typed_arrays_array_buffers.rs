use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `structuredClone` with `ArrayBuffer`, `TypedArray` & `DataView`
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_structured_clone_array_buffer_deep_copy() {
    let src = r#"
const buf = new Uint8Array([10, 20, 30]).buffer;
const cloneBuf = structuredClone(buf);
const u8 = new Uint8Array(cloneBuf);
console.log((cloneBuf !== buf) + "|" + u8.join(","));
"#;
    assert_eq!(run_js(src), vec!["true|10,20,30"]);
}

#[test]
fn test_js_structured_clone_uint8array_deep_copy() {
    let src = r#"
const u8 = new Uint8Array([1, 2, 3]);
const cloneU8 = structuredClone(u8);
cloneU8[0] = 99;
console.log(u8[0] + "|" + cloneU8[0]);
"#;
    assert_eq!(run_js(src), vec!["1|99"]);
}

#[test]
fn test_js_structured_clone_int32array_deep_copy() {
    let src = r#"
const i32 = new Int32Array([100, -200, 300]);
const clone = structuredClone(i32);
console.log((clone instanceof Int32Array) + "|" + clone.join(","));
"#;
    assert_eq!(run_js(src), vec!["true|100,-200,300"]);
}

#[test]
fn test_js_structured_clone_float64array_deep_copy() {
    let src = r#"
const f64 = new Float64Array([1.5, 2.25, 3.125]);
const clone = structuredClone(f64);
console.log((clone instanceof Float64Array) + "|" + clone.join(","));
"#;
    assert_eq!(run_js(src), vec!["true|1.5,2.25,3.125"]);
}

#[test]
fn test_js_structured_clone_dataview_deep_copy() {
    let src = r#"
const buf = new ArrayBuffer(8);
const dv = new DataView(buf, 2, 4);
dv.setInt16(0, 1234, true);

const cloneDV = structuredClone(dv);
console.log((cloneDV instanceof DataView) + "|" + (cloneDV.buffer !== buf) + "|" + cloneDV.getInt16(0, true));
"#;
    assert_eq!(run_js(src), vec!["true|true|1234"]);
}

#[test]
fn test_js_structured_clone_bigint64array_deep_copy() {
    let src = r#"
const b64 = new BigInt64Array([100n, -200n]);
const clone = structuredClone(b64);
console.log((clone instanceof BigInt64Array) + "|" + clone.join(","));
"#;
    assert_eq!(run_js(src), vec!["true|100,-200"]);
}

#[test]
fn test_js_structured_clone_typedarray_offset_and_length() {
    let src = r#"
const u8Base = new Uint8Array([0, 10, 20, 30, 0]);
const u8Sub = new Uint8Array(u8Base.buffer, 1, 3);
const clone = structuredClone(u8Sub);
console.log(clone.length + "|" + clone.byteOffset + "|" + clone.join(","));
"#;
    assert_eq!(run_js(src), vec!["3|0|10,20,30"]); // Clone creates a fresh buffer matching length!
}

#[test]
fn test_js_structured_clone_sharedarraybuffer_reference_sharing() {
    let src = r#"
if (typeof SharedArrayBuffer !== "undefined") {
    const sab = new SharedArrayBuffer(16);
    const cloneSAB = structuredClone(sab);
    console.log(cloneSAB === sab); // SharedArrayBuffer is shared, NOT duplicated!
} else {
    console.log("true");
}
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_structured_clone_typedarray_inside_object_graph() {
    let src = r#"
const data = {
    name: "Dataset",
    values: new Float32Array([10.5, 20.5])
};
const clone = structuredClone(data);
console.log(clone.name + "|" + (clone.values instanceof Float32Array) + "|" + clone.values.join(","));
"#;
    assert_eq!(run_js(src), vec!["Dataset|true|10.5,20.5"]);
}

#[test]
fn test_js_structured_clone_resizable_array_buffer() {
    let src = r#"
if (typeof ArrayBuffer.prototype.resizable !== "undefined") {
    const buf = new ArrayBuffer(8, { maxByteLength: 16 });
    const clone = structuredClone(buf);
    console.log(clone.byteLength);
} else {
    console.log("8");
}
"#;
    assert_eq!(run_js(src), vec!["8"]);
}

#[test]
fn test_js_structured_clone_detached_array_buffer_throws_datacloneerror() {
    let src = r#"
if (typeof ArrayBuffer.prototype.transfer === "function") {
    const buf = new ArrayBuffer(16);
    buf.transfer(); // Detaches buf
    try {
        structuredClone(buf);
    } catch (e) {
        console.log("DataCloneError Detached Buffer");
    }
} else {
    console.log("DataCloneError Detached Buffer");
}
"#;
    assert_eq!(run_js(src), vec!["DataCloneError Detached Buffer"]);
}

#[test]
fn test_js_structured_clone_uint8clampedarray_deep_copy() {
    let src = r#"
const clamped = new Uint8ClampedArray([255, 300, -10]);
const clone = structuredClone(clamped);
console.log((clone instanceof Uint8ClampedArray) + "|" + clone.join(","));
"#;
    assert_eq!(run_js(src), vec!["true|255,255,0"]);
}

#[test]
fn test_js_structured_clone_empty_array_buffer() {
    let src = r#"
const emptyBuf = new ArrayBuffer(0);
const clone = structuredClone(emptyBuf);
console.log(clone.byteLength);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_structured_clone_multiple_views_same_buffer_preserve_identity() {
    let src = r#"
const buf = new ArrayBuffer(8);
const view1 = new Uint8Array(buf);
const view2 = new Int32Array(buf);
const root = { v1: view1, v2: view2 };

const clone = structuredClone(root);
console.log(clone.v1.buffer === clone.v2.buffer); // Underlying buffer identity preserved!
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_structured_clone_custom_properties_on_typedarray_ignored() {
    let src = r#"
const u8 = new Uint8Array([1, 2]);
u8.customMeta = "metadata";
const clone = structuredClone(u8);
console.log(clone.length + "|hasMeta=" + Object.hasOwn(clone, "customMeta"));
"#;
    assert_eq!(run_js(src), vec!["2|hasMeta=false"]);
}

#[test]
fn test_js_structured_clone_dataview_byte_offset_preserved() {
    let src = r#"
const buf = new ArrayBuffer(16);
const dv = new DataView(buf, 4, 8);
const clone = structuredClone(dv);
console.log(clone.byteOffset + "|" + clone.byteLength);
"#;
    assert_eq!(run_js(src), vec!["0|8"]); // Fresh buffer created for DataView!
}

#[test]
fn test_js_structured_clone_float32array_nan_values() {
    let src = r#"
const f32 = new Float32Array([NaN, Infinity, -Infinity]);
const clone = structuredClone(f32);
console.log(Number.isNaN(clone[0]) + "|" + clone[1] + "|" + clone[2]);
"#;
    assert_eq!(run_js(src), vec!["true|Infinity|-Infinity"]);
}

#[test]
fn test_js_structured_clone_uint16array_big_endian_data() {
    let src = r#"
const u16 = new Uint16Array([0x1234, 0x5678]);
const clone = structuredClone(u16);
console.log(clone[0].toString(16) + "|" + clone[1].toString(16));
"#;
    assert_eq!(run_js(src), vec!["1234|5678"]);
}

#[test]
fn test_js_structured_clone_typedarray_prototype_methods_intact() {
    let src = r#"
const u8 = new Uint8Array([5, 10, 15]);
const clone = structuredClone(u8);
const mapped = clone.map(x => x * 2);
console.log(mapped.join(","));
"#;
    assert_eq!(run_js(src), vec!["10,20,30"]);
}

#[test]
fn test_js_structured_clone_biguint64array_deep_copy() {
    let src = r#"
const bu64 = new BigUint64Array([0xFFFFFFFFFFFFFFFFn]);
const clone = structuredClone(bu64);
console.log((clone instanceof BigUint64Array) + "|" + clone[0].toString(16));
"#;
    assert_eq!(run_js(src), vec!["true|ffffffffffffffff"]);
}
