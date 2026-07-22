use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: ArrayBuffer (`slice`, `transfer` ES2024, Resizable Buffers ES2024)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_arraybuffer_byte_length_property() {
    let src = r#"
const buf = new ArrayBuffer(64);
console.log(buf.byteLength);
"#;
    assert_eq!(run_js(src), vec!["64"]);
}

#[test]
fn test_js_arraybuffer_slice_copy_range() {
    let src = r#"
const buf = new ArrayBuffer(16);
const view1 = new Uint8Array(buf);
view1[4] = 99;

const slicedBuf = buf.slice(4, 8);
const view2 = new Uint8Array(slicedBuf);
console.log(slicedBuf.byteLength + "|" + view2[0] + "|isCopy=" + (slicedBuf !== buf));
"#;
    assert_eq!(run_js(src), vec!["4|99|isCopy=true"]);
}

#[test]
fn test_js_arraybuffer_slice_negative_indices() {
    let src = r#"
const buf = new ArrayBuffer(10);
const sliced = buf.slice(-4, -1);
console.log(sliced.byteLength);
"#;
    assert_eq!(run_js(src), vec!["3"]);
}

#[test]
fn test_js_arraybuffer_is_view_static_utility() {
    let src = r#"
const buf = new ArrayBuffer(16);
const u8 = new Uint8Array(buf);
const dv = new DataView(buf);

console.log(`${ArrayBuffer.isView(u8)}|${ArrayBuffer.isView(dv)}|${ArrayBuffer.isView(buf)}|${ArrayBuffer.isView({})}`);
"#;
    assert_eq!(run_js(src), vec!["true|true|false|false"]);
}

#[test]
fn test_js_arraybuffer_detached_buffer_state() {
    let src = r#"
const buf = new ArrayBuffer(8);
console.log(buf.detached);
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_arraybuffer_transfer_ownership_es2024() {
    let src = r#"
const buf = new ArrayBuffer(16);
const view = new Uint8Array(buf);
view[0] = 42;

const transferred = buf.transfer();
console.log(transferred.byteLength + "|" + new Uint8Array(transferred)[0] + "|oldDetached=" + buf.detached);
"#;
    assert_eq!(run_js(src), vec!["16|42|oldDetached=true"]);
}

#[test]
fn test_js_arraybuffer_transfer_to_smaller_size() {
    let src = r#"
const buf = new ArrayBuffer(16);
const transferred = buf.transfer(8);
console.log(transferred.byteLength + "|oldDetached=" + buf.detached);
"#;
    assert_eq!(run_js(src), vec!["8|oldDetached=true"]);
}

#[test]
fn test_js_arraybuffer_transfer_to_zero_size_detached() {
    let src = r#"
const buf = new ArrayBuffer(16);
const transferred = buf.transfer(0);
console.log(transferred.byteLength + "|detached=" + buf.detached);
"#;
    assert_eq!(run_js(src), vec!["0|detached=true"]);
}

#[test]
fn test_js_arraybuffer_transfertofixedlength_es2024() {
    let src = r#"
const buf = new ArrayBuffer(16, { maxByteLength: 32 });
const fixed = buf.transferToFixedLength(8);
console.log(fixed.resizable + "|" + fixed.byteLength);
"#;
    assert_eq!(run_js(src), vec!["false|8"]);
}

#[test]
fn test_js_arraybuffer_resizable_max_byte_length_es2024() {
    let src = r#"
const buf = new ArrayBuffer(8, { maxByteLength: 64 });
console.log(buf.resizable + "|" + buf.maxByteLength);
"#;
    assert_eq!(run_js(src), vec!["true|64"]);
}

#[test]
fn test_js_arraybuffer_resize_grow_and_shrink_es2024() {
    let src = r#"
const buf = new ArrayBuffer(8, { maxByteLength: 32 });
buf.resize(16);
console.log(buf.byteLength);
buf.resize(4);
console.log(buf.byteLength);
"#;
    assert_eq!(run_js(src), vec!["16", "4"]);
}

#[test]
fn test_js_arraybuffer_resize_beyond_max_byte_length_throws_rangeerror() {
    let src = r#"
const buf = new ArrayBuffer(8, { maxByteLength: 16 });
try {
    buf.resize(32);
} catch (e) {
    console.log("Resize Exceeds MaxByteLength RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["Resize Exceeds MaxByteLength RangeError"]);
}

#[test]
fn test_js_arraybuffer_resize_non_resizable_buffer_throws_typeerror() {
    let src = r#"
const buf = new ArrayBuffer(8);
try {
    buf.resize(16);
} catch (e) {
    console.log("Resize Fixed Buffer TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Resize Fixed Buffer TypeError"]);
}

#[test]
fn test_js_arraybuffer_detached_buffer_access_throws_typeerror() {
    let src = r#"
const buf = new ArrayBuffer(16);
const u8 = new Uint8Array(buf);
buf.transfer(); // Detaches buf

try {
    u8[0] = 10;
} catch (e) {
    console.log("Detached Buffer Access TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Detached Buffer Access TypeError"]);
}

#[test]
fn test_js_arraybuffer_slice_on_detached_buffer_throws_typeerror() {
    let src = r#"
const buf = new ArrayBuffer(16);
buf.transfer();
try {
    buf.slice(0, 4);
} catch (e) {
    console.log("Slice Detached Buffer TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Slice Detached Buffer TypeError"]);
}

#[test]
fn test_js_arraybuffer_constructor_length_coercion() {
    let src = r#"
const buf = new ArrayBuffer("32");
console.log(buf.byteLength);
"#;
    assert_eq!(run_js(src), vec!["32"]);
}

#[test]
fn test_js_arraybuffer_constructor_negative_length_throws_rangeerror() {
    let src = r#"
try {
    new ArrayBuffer(-10);
} catch (e) {
    console.log("Negative Buffer Length RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["Negative Buffer Length RangeError"]);
}

#[test]
fn test_js_arraybuffer_slice_species_constructor() {
    let src = r#"
class CustomBuffer extends ArrayBuffer {}
const buf = new CustomBuffer(16);
const sliced = buf.slice(0, 8);
console.log(sliced.byteLength + "|isCustom=" + (sliced instanceof CustomBuffer));
"#;
    assert_eq!(run_js(src), vec!["8|isCustom=true"]);
}

#[test]
fn test_js_arraybuffer_resizable_view_auto_tracking() {
    let src = r#"
const buf = new ArrayBuffer(8, { maxByteLength: 32 });
const view = new Int32Array(buf);
console.log(view.length); // 8 bytes / 4 = 2 elements
buf.resize(16);
console.log(view.length); // Resized buffer automatically updates view.length to 4 elements!
"#;
    assert_eq!(run_js(src), vec!["2", "4"]);
}

#[test]
fn test_js_arraybuffer_max_byte_length_non_resizable_defaults_to_byte_length() {
    let src = r#"
const buf = new ArrayBuffer(16);
console.log(buf.resizable + "|" + buf.maxByteLength);
"#;
    assert_eq!(run_js(src), vec!["false|16"]);
}
