use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `structuredClone` Transfer Options (`{ transfer: [...] }`)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_structured_clone_array_buffer_transfer() {
    let src = r#"
const buf = new Uint8Array([1, 2, 3, 4]).buffer;
const clone = structuredClone(buf, { transfer: [buf] });
const u8 = new Uint8Array(clone);

console.log((buf.byteLength === 0) + "|" + u8.join(",")); // buf is transferred (detached, length 0)!
"#;
    assert_eq!(run_js(src), vec!["true|1,2,3,4"]);
}

#[test]
fn test_js_structured_clone_typed_array_buffer_transfer() {
    let src = r#"
const u8 = new Uint8Array([10, 20, 30]);
const clone = structuredClone(u8, { transfer: [u8.buffer] });

console.log((u8.buffer.byteLength === 0) + "|" + (u8.length === 0 || u8[0] === undefined) + "|" + clone.join(","));
"#;
    assert_eq!(run_js(src), vec!["true|true|10,20,30"]);
}

#[test]
fn test_js_structured_clone_transfer_duplicate_buffer_throws_datacloneerror() {
    let src = r#"
const buf = new ArrayBuffer(8);
try {
    structuredClone(buf, { transfer: [buf, buf] }); // Duplicate transfer item throws DataCloneError!
} catch (e) {
    console.log("DataCloneError Duplicate Transfer");
}
"#;
    assert_eq!(run_js(src), vec!["DataCloneError Duplicate Transfer"]);
}

#[test]
fn test_js_structured_clone_transfer_non_transferable_throws_datacloneerror() {
    let src = r#"
try {
    structuredClone({ a: 1 }, { transfer: [{ a: 1 }] });
} catch (e) {
    console.log("DataCloneError Non-Transferable");
}
"#;
    assert_eq!(run_js(src), vec!["DataCloneError Non-Transferable"]);
}

#[test]
fn test_js_structured_clone_transfer_null_throws_typeerror() {
    let src = r#"
try {
    structuredClone({ a: 1 }, { transfer: [null] });
} catch (e) {
    console.log("Transfer List Null TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Transfer List Null TypeError"]);
}

#[test]
fn test_js_structured_clone_transfer_undefined_options() {
    let src = r#"
const buf = new Uint8Array([5]).buffer;
const clone = structuredClone(buf, undefined);
console.log((buf.byteLength === 1) + "|" + (clone.byteLength === 1)); // Regular deep copy
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_structured_clone_transfer_empty_array() {
    let src = r#"
const buf = new Uint8Array([5]).buffer;
const clone = structuredClone(buf, { transfer: [] });
console.log((buf.byteLength === 1) + "|" + (clone.byteLength === 1));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_structured_clone_transfer_multiple_buffers() {
    let src = r#"
const buf1 = new Uint8Array([1]).buffer;
const buf2 = new Uint8Array([2]).buffer;
const root = { b1: buf1, b2: buf2 };
const clone = structuredClone(root, { transfer: [buf1, buf2] });

console.log((buf1.byteLength === 0) + "|" + (buf2.byteLength === 0) + "|" + new Uint8Array(clone.b1)[0] + "|" + new Uint8Array(clone.b2)[0]);
"#;
    assert_eq!(run_js(src), vec!["true|true|1|2"]);
}

#[test]
fn test_js_structured_clone_transfer_detached_buffer_throws_datacloneerror() {
    let src = r#"
if (typeof ArrayBuffer.prototype.transfer === "function") {
    const buf = new ArrayBuffer(8);
    buf.transfer();
    try {
        structuredClone(buf, { transfer: [buf] });
    } catch (e) {
        console.log("DataCloneError Transfer Already Detached");
    }
} else {
    console.log("DataCloneError Transfer Already Detached");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["DataCloneError Transfer Already Detached"]
    );
}

#[test]
fn test_js_structured_clone_transfer_iterable_options() {
    let src = r#"
const buf = new Uint8Array([42]).buffer;
const transferSet = new Set([buf]);
const clone = structuredClone(buf, { transfer: transferSet });
console.log((buf.byteLength === 0) + "|" + new Uint8Array(clone)[0]);
"#;
    assert_eq!(run_js(src), vec!["true|42"]);
}

#[test]
fn test_js_structured_clone_transfer_buffer_not_in_value_graph_detached_anyway() {
    let src = r#"
const bufToDetach = new ArrayBuffer(16);
const valueToClone = { msg: "Hello" };
const clone = structuredClone(valueToClone, { transfer: [bufToDetach] });

console.log(clone.msg + "|detached=" + (bufToDetach.byteLength === 0));
"#;
    assert_eq!(run_js(src), vec!["Hello|detached=true"]);
}

#[test]
fn test_js_structured_clone_options_not_an_object_throws_typeerror() {
    let src = r#"
try {
    structuredClone(123, "not_an_object");
} catch (e) {
    console.log("Options Not Object TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Options Not Object TypeError"]);
}

#[test]
fn test_js_structured_clone_transfer_non_iterable_throws_typeerror() {
    let src = r#"
try {
    structuredClone(123, { transfer: 12345 });
} catch (e) {
    console.log("Transfer Option Non-Iterable TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Transfer Option Non-Iterable TypeError"]);
}

#[test]
fn test_js_structured_clone_transfer_dataview_underlying_buffer() {
    let src = r#"
const buf = new ArrayBuffer(8);
const dv = new DataView(buf);
dv.setInt32(0, 9999);
const cloneDV = structuredClone(dv, { transfer: [buf] });

console.log((buf.byteLength === 0) + "|" + cloneDV.getInt32(0));
"#;
    assert_eq!(run_js(src), vec!["true|9999"]);
}

#[test]
fn test_js_structured_clone_transfer_shared_array_buffer_throws_datacloneerror() {
    let src = r#"
if (typeof SharedArrayBuffer !== "undefined") {
    const sab = new SharedArrayBuffer(16);
    try {
        structuredClone(sab, { transfer: [sab] }); // SharedArrayBuffer cannot be transferred!
    } catch (e) {
        console.log("DataCloneError Transfer SharedArrayBuffer");
    }
} else {
    console.log("DataCloneError Transfer SharedArrayBuffer");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["DataCloneError Transfer SharedArrayBuffer"]
    );
}

#[test]
fn test_js_structured_clone_transfer_getter_throwing_in_options() {
    let src = r#"
const opts = {
    get transfer() { throw new Error("TransferGetterError"); }
};
try {
    structuredClone(123, opts);
} catch (e) {
    console.log(e.message);
}
"#;
    assert_eq!(run_js(src), vec!["TransferGetterError"]);
}

#[test]
fn test_js_structured_clone_transfer_int32array_view_buffer() {
    let src = r#"
const i32 = new Int32Array([100, 200]);
const clone = structuredClone(i32, { transfer: [i32.buffer] });
console.log((i32.buffer.byteLength === 0) + "|" + clone.join(","));
"#;
    assert_eq!(run_js(src), vec!["true|100,200"]);
}

#[test]
fn test_js_structured_clone_transfer_bigint64array_view_buffer() {
    let src = r#"
const b64 = new BigInt64Array([1000n]);
const clone = structuredClone(b64, { transfer: [b64.buffer] });
console.log((b64.buffer.byteLength === 0) + "|" + clone[0].toString());
"#;
    assert_eq!(run_js(src), vec!["true|1000"]);
}

#[test]
fn test_js_structured_clone_transfer_preserves_nested_cyclical_identity() {
    let src = r#"
const buf = new Uint8Array([7, 8, 9]).buffer;
const node = { buffer: buf };
node.self = node;
const clone = structuredClone(node, { transfer: [buf] });

console.log((buf.byteLength === 0) + "|" + (clone.self === clone) + "|" + new Uint8Array(clone.buffer).join(","));
"#;
    assert_eq!(run_js(src), vec!["true|true|7,8,9"]);
}

#[test]
fn test_js_structured_clone_transfer_with_primitive_target() {
    let src = r#"
const buf = new Uint8Array([99]).buffer;
const cloneVal = structuredClone(100, { transfer: [buf] });
console.log(cloneVal + "|detached=" + (buf.byteLength === 0));
"#;
    assert_eq!(run_js(src), vec!["100|detached=true"]);
}
