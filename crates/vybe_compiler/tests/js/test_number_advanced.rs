/// Number and math advanced — precision, formatting, number theory

use super::helpers::run_js;

#[test]
fn integer_overflow_bigint() {
    assert_eq!(run_js(r#"
const MAX_SAFE = Number.MAX_SAFE_INTEGER;
console.log(MAX_SAFE + 1 === MAX_SAFE + 2);  // loses precision
const bigMax = BigInt(MAX_SAFE);
console.log(bigMax + 1n === bigMax + 2n);    // precise
console.log((bigMax + 1n).toString());
"#), vec!["true", "false", "9007199254740992"]);
}

#[test]
fn float_comparison_epsilon() {
    assert_eq!(run_js(r#"
function aboutEqual(a, b, eps = Number.EPSILON) {
    return Math.abs(a - b) <= eps * Math.max(Math.abs(a), Math.abs(b));
}
console.log(0.1 + 0.2 === 0.3);
console.log(aboutEqual(0.1 + 0.2, 0.3, 1e-10));
console.log(aboutEqual(1.0, 1.0 + 1e-16, 1e-10));
"#), vec!["false", "true", "true"]);
}

#[test]
fn integer_division_floor_trunc_sign() {
    assert_eq!(run_js(r#"
// Floor division (Python-style)
const floorDiv = (a, b) => Math.floor(a / b);
// Truncation division (C-style)
const truncDiv = (a, b) => Math.trunc(a / b);
console.log(floorDiv(7, 2));
console.log(floorDiv(-7, 2));
console.log(truncDiv(-7, 2));
console.log(floorDiv(7, -2));
"#), vec!["3", "-4", "-3", "-4"]);
}

#[test]
fn bit_manipulation() {
    assert_eq!(run_js(r#"
// Set bit
const set = (n, bit) => n | (1 << bit);
// Clear bit
const clear = (n, bit) => n & ~(1 << bit);
// Toggle bit
const toggle = (n, bit) => n ^ (1 << bit);
// Check bit
const check = (n, bit) => !!(n & (1 << bit));
let n = 0;
n = set(n, 0); n = set(n, 2); n = set(n, 4);
console.log(n.toString(2));
console.log(check(n, 2));
n = clear(n, 2);
console.log(check(n, 2));
console.log(toggle(n, 0).toString(2));
"#), vec!["10101", "true", "false", "10100"]);
}

#[test]
fn number_rounding_modes() {
    assert_eq!(run_js(r#"
// Round half up (standard)
const roundHalfUp = n => Math.floor(n + 0.5);
// Round half to even (banker's rounding simulation)
const roundHalfEven = n => {
    const floor = Math.floor(n);
    const frac = n - floor;
    if (Math.abs(frac - 0.5) < 1e-10) {
        return floor % 2 === 0 ? floor : floor + 1;
    }
    return Math.round(n);
};
console.log(roundHalfUp(2.5));
console.log(roundHalfUp(-2.5));
console.log(roundHalfEven(2.5));
console.log(roundHalfEven(3.5));
"#), vec!["3", "-2", "2", "4"]);
}

#[test]
fn modular_arithmetic() {
    assert_eq!(run_js(r#"
// JS % can be negative — true modulo
const mod = (a, n) => ((a % n) + n) % n;
console.log(mod(7, 3));
console.log(mod(-7, 3));
console.log(mod(100, 7));
"#), vec!["1", "2", "2"]);
}

#[test]
fn numeric_types_detection() {
    assert_eq!(run_js(r#"
const isInt = n => Number.isInteger(n);
const isFloat = n => typeof n === "number" && !Number.isInteger(n) && isFinite(n);
const isBigInt = n => typeof n === "bigint";
console.log(isInt(5));
console.log(isInt(5.0));
console.log(isFloat(5.5));
console.log(isFloat(Infinity));
console.log(isBigInt(42n));
"#), vec!["true", "true", "true", "false", "true"]);
}

#[test]
fn hex_octal_binary_ops() {
    assert_eq!(run_js(r#"
const hex = 0xFF;
const octal = 0o17;
const binary = 0b1010;
console.log(hex);
console.log(octal);
console.log(binary);
console.log((hex & binary).toString(2));
console.log((octal | binary).toString(2));
"#), vec!["255", "15", "10", "1010", "1111"]);
}

#[test]
fn number_to_string_bases() {
    assert_eq!(run_js(r#"
const n = 255;
console.log(n.toString(2));
console.log(n.toString(8));
console.log(n.toString(16));
console.log(n.toString(36));
// And parsing back
console.log(parseInt("ff", 16));
console.log(parseInt("11111111", 2));
"#), vec!["11111111", "377", "ff", "73", "255", "255"]);
}

#[test]
fn math_advanced_functions() {
    assert_eq!(run_js(r#"
console.log(Math.sign(-5));
console.log(Math.sign(0));
console.log(Math.sign(7));
console.log(Math.hypot(3, 4));
console.log(Math.log2(8));
console.log(Math.log10(1000));
"#), vec!["-1", "0", "1", "5", "3", "3"]);
}

#[test]
fn number_parsing_edge_cases() {
    assert_eq!(run_js(r#"
console.log(parseInt("0x1F"));
console.log(parseInt("077"));
console.log(parseInt("3.99"));
console.log(parseFloat("3.14abc"));
console.log(Number("  42  "));
console.log(Number(""));
"#), vec!["31", "77", "3", "3.14", "42", "0"]);
}

#[test]
fn random_seeded_lcg() {
    assert_eq!(run_js(r#"
// Simple LCG pseudo-random for deterministic tests
function lcg(seed) {
    let s = seed;
    return () => {
        s = (1664525 * s + 1013904223) & 0xFFFFFFFF;
        return (s >>> 0) / 0x100000000;
    };
}
const rand = lcg(42);
const vals = Array.from({length: 5}, () => rand() > 0 && rand() < 1);
console.log(vals.every(Boolean));
const r1 = lcg(42)();
const r2 = lcg(42)();
console.log(r1 === r2);
"#), vec!["true", "true"]);
}
