use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Atomics.isLockFree()` & Lock-Free Hardware Introspection
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_atomics_is_lock_free_one_byte() {
    let src = r#"
console.log(typeof Atomics.isLockFree(1) === "boolean");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_atomics_is_lock_free_two_bytes() {
    let src = r#"
console.log(typeof Atomics.isLockFree(2) === "boolean");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_atomics_is_lock_free_four_bytes() {
    let src = r#"
console.log(Atomics.isLockFree(4));
"#;
    assert_eq!(run_js(src), vec!["true"]); // 4-byte atomic operations are ALWAYS lock-free on modern architectures!
}

#[test]
fn test_js_atomics_is_lock_free_eight_bytes() {
    let src = r#"
console.log(typeof Atomics.isLockFree(8) === "boolean");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_atomics_is_lock_free_unsupported_sizes_return_false() {
    let src = r#"
console.log(`${Atomics.isLockFree(3)}:${Atomics.isLockFree(5)}:${Atomics.isLockFree(7)}:${Atomics.isLockFree(9)}`);
"#;
    assert_eq!(run_js(src), vec!["false:false:false:false"]); // Non-power-of-2 sizes (3, 5, 7, 9) return false!
}

#[test]
fn test_js_atomics_is_lock_free_zero_returns_false() {
    let src = r#"
console.log(Atomics.isLockFree(0));
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_atomics_is_lock_free_negative_returns_false() {
    let src = r#"
console.log(Atomics.isLockFree(-4));
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_atomics_is_lock_free_coerces_arg_to_integer() {
    let src = r#"
console.log(Atomics.isLockFree("4.9"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_atomics_is_lock_free_nan_returns_false() {
    let src = r#"
console.log(Atomics.isLockFree(NaN));
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_atomics_is_lock_free_infinity_returns_false() {
    let src = r#"
console.log(Atomics.isLockFree(Infinity));
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_atomics_is_lock_free_property_descriptor() {
    let src = r#"
const desc = Object.getOwnPropertyDescriptor(Atomics, "isLockFree");
console.log(`${desc.writable}:${desc.enumerable}:${desc.configurable}:${Atomics.isLockFree.length}`);
"#;
    assert_eq!(run_js(src), vec!["true:false:true:1"]);
}

#[test]
fn test_js_atomics_is_lock_free_name_property() {
    let src = r#"
console.log(Atomics.isLockFree.name);
"#;
    assert_eq!(run_js(src), vec!["isLockFree"]);
}

#[test]
fn test_js_atomics_is_lock_free_typed_array_element_bytes() {
    let src = r#"
console.log([
    Atomics.isLockFree(Int8Array.BYTES_PER_ELEMENT),
    Atomics.isLockFree(Int16Array.BYTES_PER_ELEMENT),
    Atomics.isLockFree(Int32Array.BYTES_PER_ELEMENT),
    Atomics.isLockFree(BigInt64Array.BYTES_PER_ELEMENT)
].every(res => typeof res === "boolean"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_atomics_is_lock_free_without_arguments_returns_false() {
    let src = r#"
console.log(Atomics.isLockFree());
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_atomics_is_lock_free_boolean_arg_coerced_to_number() {
    let src = r#"
console.log(Atomics.isLockFree(true)); // true coerces to 1
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_atomics_is_lock_free_null_arg_coerced_to_zero() {
    let src = r#"
console.log(Atomics.isLockFree(null)); // null coerces to 0
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_atomics_is_lock_free_object_valueof_coercion() {
    let src = r#"
const obj = { valueOf: () => 4 };
console.log(Atomics.isLockFree(obj));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_atomics_is_lock_free_symbol_throws_typeerror() {
    let src = r#"
try {
    Atomics.isLockFree(Symbol("4"));
} catch (e) {
    console.log("isLockFree Symbol TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["isLockFree Symbol TypeError"]);
}

#[test]
fn test_js_atomics_is_lock_free_large_power_of_two_returns_false() {
    let src = r#"
console.log(Atomics.isLockFree(16) + "|" + Atomics.isLockFree(32));
"#;
    assert_eq!(run_js(src), vec!["false|false"]);
}

#[test]
fn test_js_atomics_is_lock_free_deterministic_across_calls() {
    let src = r#"
const r1 = Atomics.isLockFree(4);
const r2 = Atomics.isLockFree(4);
console.log(r1 === r2);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}
