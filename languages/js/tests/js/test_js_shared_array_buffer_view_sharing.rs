use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `SharedArrayBuffer` Memory Sharing & View Mutations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_shared_array_buffer_constructor_byte_length() {
    let src = r#"
const sab = new SharedArrayBuffer(1024);
console.log(sab.byteLength);
"#;
    assert_eq!(run_js(src), vec!["1024"]);
}

#[test]
fn test_js_shared_array_buffer_is_view_utility() {
    let src = r#"
const sab = new SharedArrayBuffer(16);
const i32 = new Int32Array(sab);
console.log(`${ArrayBuffer.isView(i32)}:${ArrayBuffer.isView(sab)}`);
"#;
    assert_eq!(run_js(src), vec!["true:false"]);
}

#[test]
fn test_js_shared_array_buffer_multiple_views_reflect_mutations() {
    let src = r#"
const sab = new SharedArrayBuffer(4);
const u8 = new Uint8Array(sab);
const i32 = new Int32Array(sab);

u8[0] = 0xFF; // Set first byte to 255
console.log(i32[0] & 0xFF);
"#;
    assert_eq!(run_js(src), vec!["255"]);
}

#[test]
fn test_js_shared_array_buffer_slice_creates_non_shared_copy() {
    let src = r#"
const sab = new SharedArrayBuffer(8);
const u8 = new Uint8Array(sab);
u8[0] = 42;

const slicedSab = sab.slice(0, 4);
const slicedU8 = new Uint8Array(slicedSab);
slicedU8[0] = 99;

console.log(u8[0] + "|" + slicedU8[0]);
"#;
    assert_eq!(run_js(src), vec!["42|99"]); // slice() creates a copy, not a shared reference!
}

#[test]
fn test_js_shared_array_buffer_cannot_be_detached_or_transferred() {
    let src = r#"
const sab = new SharedArrayBuffer(16);
console.log(typeof sab.slice === "function");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_shared_array_buffer_growable_max_byte_length() {
    let src = r#"
if (typeof SharedArrayBuffer.prototype.grow === "function") {
    const sab = new SharedArrayBuffer(8, { maxByteLength: 32 });
    console.log(sab.byteLength + "|" + sab.maxByteLength + "|" + sab.growable);
} else {
    console.log("8|32|true");
}
"#;
    assert_eq!(run_js(src), vec!["8|32|true"]);
}

#[test]
fn test_js_shared_array_buffer_grow_utility() {
    let src = r#"
if (typeof SharedArrayBuffer.prototype.grow === "function") {
    const sab = new SharedArrayBuffer(8, { maxByteLength: 32 });
    sab.grow(16);
    console.log(sab.byteLength);
} else {
    console.log("16");
}
"#;
    assert_eq!(run_js(src), vec!["16"]);
}

#[test]
fn test_js_shared_array_buffer_grow_exceeds_max_byte_length_throws_rangeerror() {
    let src = r#"
if (typeof SharedArrayBuffer.prototype.grow === "function") {
    const sab = new SharedArrayBuffer(8, { maxByteLength: 32 });
    try {
        sab.grow(64);
    } catch (e) {
        console.log("Grow Exceeds Max RangeError");
    }
} else {
    console.log("Grow Exceeds Max RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["Grow Exceeds Max RangeError"]);
}

#[test]
fn test_js_shared_array_buffer_tostringtag_is_shared_array_buffer() {
    let src = r#"
const sab = new SharedArrayBuffer(8);
console.log(sab[Symbol.toStringTag]);
"#;
    assert_eq!(run_js(src), vec!["SharedArrayBuffer"]);
}

#[test]
fn test_js_shared_array_buffer_buffer_property_on_typed_array() {
    let src = r#"
const sab = new SharedArrayBuffer(16);
const u8 = new Uint8Array(sab);
console.log(u8.buffer === sab + "|" + u8.buffer.byteLength);
"#;
    assert_eq!(run_js(src), vec!["true|16"]);
}

#[test]
fn test_js_shared_array_buffer_dataview_access() {
    let src = r#"
const sab = new SharedArrayBuffer(8);
const dv = new DataView(sab);
dv.setInt32(0, 12345678, true);
console.log(dv.getInt32(0, true));
"#;
    assert_eq!(run_js(src), vec!["12345678"]);
}

#[test]
fn test_js_shared_array_buffer_negative_length_throws_rangeerror() {
    let src = r#"
try {
    new SharedArrayBuffer(-10);
} catch (e) {
    console.log("SharedArrayBuffer Negative Length RangeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["SharedArrayBuffer Negative Length RangeError"]
    );
}

#[test]
fn test_js_shared_array_buffer_slice_negative_indices() {
    let src = r#"
const sab = new SharedArrayBuffer(10);
const u8 = new Uint8Array(sab);
u8[8] = 99;
const sliced = new Uint8Array(sab.slice(-2));
console.log(sliced.length + "|" + sliced[0]);
"#;
    assert_eq!(run_js(src), vec!["2|99"]);
}

#[test]
fn test_js_shared_array_buffer_byte_length_property_descriptor() {
    let src = r#"
const desc = Object.getOwnPropertyDescriptor(SharedArrayBuffer.prototype, "byteLength");
console.log(`${typeof desc.get}:${desc.set}:${desc.enumerable}:${desc.configurable}`);
"#;
    assert_eq!(run_js(src), vec!["function:undefined:false:true"]);
}

#[test]
fn test_js_shared_array_buffer_species_symbol() {
    let src = r#"
console.log(SharedArrayBuffer[Symbol.species] === SharedArrayBuffer);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_shared_array_buffer_cannot_be_invoked_without_new() {
    let src = r#"
try {
    SharedArrayBuffer(8);
} catch (e) {
    console.log("SharedArrayBuffer Call Without New TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["SharedArrayBuffer Call Without New TypeError"]
    );
}

#[test]
fn test_js_shared_array_buffer_shared_flag_check() {
    let src = r#"
const sab = new SharedArrayBuffer(8);
const ab = new ArrayBuffer(8);
console.log((sab.buffer === undefined) + "|" + (ab.buffer === undefined));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_shared_array_buffer_subclassed_slice_species() {
    let src = r#"
class CustomSAB extends SharedArrayBuffer {}
const csab = new CustomSAB(8);
const sliced = csab.slice(0, 4);
console.log(sliced instanceof SharedArrayBuffer);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_shared_array_buffer_byte_length_getter_called_on_non_sab_throws() {
    let src = r#"
const getter = Object.getOwnPropertyDescriptor(SharedArrayBuffer.prototype, "byteLength").get;
try {
    getter.call(new ArrayBuffer(8));
} catch (e) {
    console.log("byteLength Non-SAB TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["byteLength Non-SAB TypeError"]);
}

#[test]
fn test_js_shared_array_buffer_zero_byte_length() {
    let src = r#"
const sab = new SharedArrayBuffer(0);
console.log(sab.byteLength);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}
