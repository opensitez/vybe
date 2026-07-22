use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Atomics.compareExchange()`, `Atomics.store()`, `Atomics.load()`
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_atomics_load_reads_value() {
    let src = r#"
const sab = new SharedArrayBuffer(4);
const i32 = new Int32Array(sab);
i32[0] = 42;
console.log(Atomics.load(i32, 0));
"#;
    assert_eq!(run_js(src), vec!["42"]);
}

#[test]
fn test_js_atomics_store_writes_value_and_returns_stored_value() {
    let src = r#"
const sab = new SharedArrayBuffer(4);
const i32 = new Int32Array(sab);
const stored = Atomics.store(i32, 0, 99);
console.log(stored + "|" + Atomics.load(i32, 0));
"#;
    assert_eq!(run_js(src), vec!["99|99"]);
}

#[test]
fn test_js_atomics_compare_exchange_success_when_expected_matches() {
    let src = r#"
const sab = new SharedArrayBuffer(4);
const i32 = new Int32Array(sab);
i32[0] = 10;
const old = Atomics.compareExchange(i32, 0, 10, 20); // expected 10 == current 10 -> swaps to 20!
console.log(old + "|" + i32[0]);
"#;
    assert_eq!(run_js(src), vec!["10|20"]);
}

#[test]
fn test_js_atomics_compare_exchange_noop_when_expected_mismatches() {
    let src = r#"
const sab = new SharedArrayBuffer(4);
const i32 = new Int32Array(sab);
i32[0] = 10;
const old = Atomics.compareExchange(i32, 0, 99, 20); // expected 99 != current 10 -> no swap!
console.log(old + "|" + i32[0]);
"#;
    assert_eq!(run_js(src), vec!["10|10"]);
}

#[test]
fn test_js_atomics_load_store_bigint64() {
    let src = r#"
const sab = new SharedArrayBuffer(8);
const bi64 = new BigInt64Array(sab);
Atomics.store(bi64, 0, 123456789n);
console.log(Atomics.load(bi64, 0).toString());
"#;
    assert_eq!(run_js(src), vec!["123456789"]);
}

#[test]
fn test_js_atomics_compare_exchange_bigint64() {
    let src = r#"
const sab = new SharedArrayBuffer(8);
const bi64 = new BigInt64Array(sab);
bi64[0] = 100n;
const old = Atomics.compareExchange(bi64, 0, 100n, 500n);
console.log(old.toString() + "|" + bi64[0].toString());
"#;
    assert_eq!(run_js(src), vec!["100|500"]);
}

#[test]
fn test_js_atomics_store_coerces_value_to_typed_array_type() {
    let src = r#"
const u8 = new Uint8Array(new SharedArrayBuffer(1));
Atomics.store(u8, 0, "150");
console.log(Atomics.load(u8, 0));
"#;
    assert_eq!(run_js(src), vec!["150"]);
}

#[test]
fn test_js_atomics_load_out_of_bounds_throws_rangeerror() {
    let src = r#"
const i32 = new Int32Array(new SharedArrayBuffer(4));
try {
    Atomics.load(i32, 2);
} catch (e) {
    console.log("Atomics.load Out of Bounds RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["Atomics.load Out of Bounds RangeError"]);
}

#[test]
fn test_js_atomics_store_out_of_bounds_throws_rangeerror() {
    let src = r#"
const i32 = new Int32Array(new SharedArrayBuffer(4));
try {
    Atomics.store(i32, -1, 10);
} catch (e) {
    console.log("Atomics.store Out of Bounds RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["Atomics.store Out of Bounds RangeError"]);
}

#[test]
fn test_js_atomics_spin_lock_simulation() {
    let src = r#"
const lock = new Int32Array(new SharedArrayBuffer(4));
// Try acquire lock (0 -> 1)
const acquired = Atomics.compareExchange(lock, 0, 0, 1) === 0;
// Release lock (1 -> 0)
if (acquired) Atomics.store(lock, 0, 0);
console.log(acquired + "|" + Atomics.load(lock, 0));
"#;
    assert_eq!(run_js(src), vec!["true|0"]);
}

#[test]
fn test_js_atomics_load_float32_array_throws_typeerror() {
    let src = r#"
const f32 = new Float32Array(1);
try {
    Atomics.load(f32, 0);
} catch (e) {
    console.log("Atomics.load Float32 TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Atomics.load Float32 TypeError"]);
}

#[test]
fn test_js_atomics_store_float64_array_throws_typeerror() {
    let src = r#"
const f64 = new Float64Array(1);
try {
    Atomics.store(f64, 0, 1.5);
} catch (e) {
    console.log("Atomics.store Float64 TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Atomics.store Float64 TypeError"]);
}

#[test]
fn test_js_atomics_compare_exchange_nan_in_int32_array() {
    let src = r#"
const i32 = new Int32Array(new SharedArrayBuffer(4));
i32[0] = 0;
const old = Atomics.compareExchange(i32, 0, NaN, 100); // NaN coerces to expected 0!
console.log(old + "|" + i32[0]);
"#;
    assert_eq!(run_js(src), vec!["0|100"]);
}

#[test]
fn test_js_atomics_load_uint32_array() {
    let src = r#"
const u32 = new Uint32Array(new SharedArrayBuffer(4));
Atomics.store(u32, 0, 4294967295);
console.log(Atomics.load(u32, 0));
"#;
    assert_eq!(run_js(src), vec!["4294967295"]);
}

#[test]
fn test_js_atomics_load_int16_array() {
    let src = r#"
const i16 = new Int16Array(new SharedArrayBuffer(2));
Atomics.store(i16, 0, -32768);
console.log(Atomics.load(i16, 0));
"#;
    assert_eq!(run_js(src), vec!["-32768"]);
}

#[test]
fn test_js_atomics_load_uint16_array() {
    let src = r#"
const u16 = new Uint16Array(new SharedArrayBuffer(2));
Atomics.store(u16, 0, 65535);
console.log(Atomics.load(u16, 0));
"#;
    assert_eq!(run_js(src), vec!["65535"]);
}

#[test]
fn test_js_atomics_compare_exchange_negative_expected_value() {
    let src = r#"
const i8 = new Int8Array(new SharedArrayBuffer(1));
i8[0] = -50;
const old = Atomics.compareExchange(i8, 0, -50, 100);
console.log(old + "|" + i8[0]);
"#;
    assert_eq!(run_js(src), vec!["-50|100"]);
}

#[test]
fn test_js_atomics_load_property_descriptor() {
    let src = r#"
const desc = Object.getOwnPropertyDescriptor(Atomics, "load");
console.log(`${desc.writable}:${desc.enumerable}:${desc.configurable}:${Atomics.load.length}`);
"#;
    assert_eq!(run_js(src), vec!["true:false:true:2"]);
}

#[test]
fn test_js_atomics_store_property_descriptor() {
    let src = r#"
const desc = Object.getOwnPropertyDescriptor(Atomics, "store");
console.log(`${desc.writable}:${desc.enumerable}:${desc.configurable}:${Atomics.store.length}`);
"#;
    assert_eq!(run_js(src), vec!["true:false:true:3"]);
}

#[test]
fn test_js_atomics_compare_exchange_property_descriptor() {
    let src = r#"
const desc = Object.getOwnPropertyDescriptor(Atomics, "compareExchange");
console.log(`${desc.writable}:${desc.enumerable}:${desc.configurable}:${Atomics.compareExchange.length}`);
"#;
    assert_eq!(run_js(src), vec!["true:false:true:4"]);
}
