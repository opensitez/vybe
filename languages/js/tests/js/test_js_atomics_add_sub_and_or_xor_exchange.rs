use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Atomics` Math & Bitwise Operations (`add`, `sub`, `and`, `or`, `xor`, `exchange`)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_atomics_add_returns_old_value() {
    let src = r#"
const sab = new SharedArrayBuffer(4);
const i32 = new Int32Array(sab);
i32[0] = 10;
const old = Atomics.add(i32, 0, 5);
console.log(old + "|" + i32[0]);
"#;
    assert_eq!(run_js(src), vec!["10|15"]);
}

#[test]
fn test_js_atomics_sub_returns_old_value() {
    let src = r#"
const sab = new SharedArrayBuffer(4);
const i32 = new Int32Array(sab);
i32[0] = 20;
const old = Atomics.sub(i32, 0, 8);
console.log(old + "|" + i32[0]);
"#;
    assert_eq!(run_js(src), vec!["20|12"]);
}

#[test]
fn test_js_atomics_and_bitwise_operation() {
    let src = r#"
const sab = new SharedArrayBuffer(4);
const i32 = new Int32Array(sab);
i32[0] = 0b1111;
const old = Atomics.and(i32, 0, 0b1010);
console.log(old + "|" + i32[0]);
"#;
    assert_eq!(run_js(src), vec!["15|10"]);
}

#[test]
fn test_js_atomics_or_bitwise_operation() {
    let src = r#"
const sab = new SharedArrayBuffer(4);
const i32 = new Int32Array(sab);
i32[0] = 0b0101;
const old = Atomics.or(i32, 0, 0b1010);
console.log(old + "|" + i32[0]);
"#;
    assert_eq!(run_js(src), vec!["5|15"]);
}

#[test]
fn test_js_atomics_xor_bitwise_operation() {
    let src = r#"
const sab = new SharedArrayBuffer(4);
const i32 = new Int32Array(sab);
i32[0] = 0b1100;
const old = Atomics.xor(i32, 0, 0b1010);
console.log(old + "|" + i32[0]);
"#;
    assert_eq!(run_js(src), vec!["12|6"]);
}

#[test]
fn test_js_atomics_exchange_stores_and_returns_old_value() {
    let src = r#"
const sab = new SharedArrayBuffer(4);
const i32 = new Int32Array(sab);
i32[0] = 100;
const old = Atomics.exchange(i32, 0, 500);
console.log(old + "|" + i32[0]);
"#;
    assert_eq!(run_js(src), vec!["100|500"]);
}

#[test]
fn test_js_atomics_operations_on_non_shared_typed_array() {
    let src = r#"
const i32 = new Int32Array(1);
i32[0] = 5;
const old = Atomics.add(i32, 0, 10); // Atomics math operations work on non-shared TypedArrays as well!
console.log(old + "|" + i32[0]);
"#;
    assert_eq!(run_js(src), vec!["5|15"]);
}

#[test]
fn test_js_atomics_operations_on_float_array_throws_typeerror() {
    let src = r#"
const f32 = new Float32Array(1);
try {
    Atomics.add(f32, 0, 1);
} catch (e) {
    console.log("Atomics Float TypedArray TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Atomics Float TypedArray TypeError"]);
}

#[test]
fn test_js_atomics_index_out_of_bounds_throws_rangeerror() {
    let src = r#"
const i32 = new Int32Array(2);
try {
    Atomics.add(i32, 5, 1);
} catch (e) {
    console.log("Atomics Index Out of Bounds RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["Atomics Index Out of Bounds RangeError"]);
}

#[test]
fn test_js_atomics_bigint_typed_array() {
    let src = r#"
const sab = new SharedArrayBuffer(8);
const bi64 = new BigInt64Array(sab);
bi64[0] = 100n;
const old = Atomics.add(bi64, 0, 50n);
console.log(old.toString() + "|" + bi64[0].toString());
"#;
    assert_eq!(run_js(src), vec!["100|150"]);
}

#[test]
fn test_js_atomics_biguint64_typed_array() {
    let src = r#"
const sab = new SharedArrayBuffer(8);
const bu64 = new BigUint64Array(sab);
bu64[0] = 200n;
const old = Atomics.sub(bu64, 0, 50n);
console.log(old.toString() + "|" + bu64[0].toString());
"#;
    assert_eq!(run_js(src), vec!["200|150"]);
}

#[test]
fn test_js_atomics_add_overflow_wrap_around() {
    let src = r#"
const u8 = new Uint8Array(new SharedArrayBuffer(1));
u8[0] = 255;
Atomics.add(u8, 0, 1);
console.log(u8[0]);
"#;
    assert_eq!(run_js(src), vec!["0"]); // Uint8Array wraps 255 + 1 to 0!
}

#[test]
fn test_js_atomics_sub_underflow_wrap_around() {
    let src = r#"
const u8 = new Uint8Array(new SharedArrayBuffer(1));
u8[0] = 0;
Atomics.sub(u8, 0, 1);
console.log(u8[0]);
"#;
    assert_eq!(run_js(src), vec!["255"]);
}

#[test]
fn test_js_atomics_coerces_index_to_integer() {
    let src = r#"
const i32 = new Int32Array(new SharedArrayBuffer(8));
i32[1] = 10;
Atomics.add(i32, "1.9", 5); // Coerces "1.9" to index 1
console.log(i32[1]);
"#;
    assert_eq!(run_js(src), vec!["15"]);
}

#[test]
fn test_js_atomics_coerces_value_to_integer() {
    let src = r#"
const i32 = new Int32Array(new SharedArrayBuffer(4));
Atomics.add(i32, 0, "20");
console.log(i32[0]);
"#;
    assert_eq!(run_js(src), vec!["20"]);
}

#[test]
fn test_js_atomics_non_typed_array_target_throws_typeerror() {
    let src = r#"
try {
    Atomics.add([1, 2], 0, 1);
} catch (e) {
    console.log("Atomics Non-TypedArray TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Atomics Non-TypedArray TypeError"]);
}

#[test]
fn test_js_atomics_add_property_descriptors() {
    let src = r#"
const desc = Object.getOwnPropertyDescriptor(Atomics, "add");
console.log(`${desc.writable}:${desc.enumerable}:${desc.configurable}:${Atomics.add.length}`);
"#;
    assert_eq!(run_js(src), vec!["true:false:true:3"]);
}

#[test]
fn test_js_atomics_tostringtag_is_atomics() {
    let src = r#"
console.log(Atomics[Symbol.toStringTag]);
"#;
    assert_eq!(run_js(src), vec!["Atomics"]);
}

#[test]
fn test_js_atomics_cannot_be_constructed_with_new() {
    let src = r#"
try {
    new Atomics();
} catch (e) {
    console.log("Atomics Constructor TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Atomics Constructor TypeError"]);
}

#[test]
fn test_js_atomics_int8_typed_array() {
    let src = r#"
const i8 = new Int8Array(new SharedArrayBuffer(1));
i8[0] = 127;
Atomics.add(i8, 0, 1);
console.log(i8[0]); // Int8Array wraps 127 + 1 to -128!
"#;
    assert_eq!(run_js(src), vec!["-128"]);
}
