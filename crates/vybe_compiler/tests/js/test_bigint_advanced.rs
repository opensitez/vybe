/// BigInt advanced patterns — arithmetic, comparison, coercion, bitwise, mixed operations
use super::helpers::run_js;

#[test]
fn bigint_basic_arithmetic() {
    assert_eq!(
        run_js(
            r#"
const a = 9007199254740993n; // beyond MAX_SAFE_INTEGER
const b = 1n;
console.log((a + b).toString());
console.log((a - b).toString());
console.log((a * 2n).toString());
"#
        ),
        vec!["9007199254740994", "9007199254740992", "18014398509481986"]
    );
}

#[test]
fn bigint_division_truncates() {
    assert_eq!(
        run_js(
            r#"
console.log(7n / 2n);
console.log(-7n / 2n);
"#
        ),
        vec!["3n", "-3n"]
    );
}

#[test]
fn bigint_remainder() {
    assert_eq!(
        run_js(
            r#"
console.log(10n % 3n);
console.log(-10n % 3n);
"#
        ),
        vec!["1n", "-1n"]
    );
}

#[test]
fn bigint_exponentiation() {
    assert_eq!(
        run_js(
            r#"
console.log(2n ** 64n > 0n);
console.log((2n ** 10n).toString());
"#
        ),
        vec!["true", "1024"]
    );
}

#[test]
fn bigint_comparison_with_number() {
    assert_eq!(
        run_js(
            r#"
console.log(1n < 2);
console.log(2n > 1);
console.log(1n == 1);    // abstract equality
console.log(1n === 1);   // strict: false (different types)
"#
        ),
        vec!["true", "true", "true", "false"]
    );
}

#[test]
fn bigint_bitwise_operations() {
    assert_eq!(
        run_js(
            r#"
console.log((0b1100n & 0b1010n).toString());
console.log((0b1100n | 0b1010n).toString());
console.log((0b1100n ^ 0b1010n).toString());
console.log((~0b1100n).toString()); // two's complement
"#
        ),
        vec!["8", "14", "6", "-13"]
    );
}

#[test]
fn bigint_shift() {
    assert_eq!(
        run_js(
            r#"
console.log((1n << 4n).toString());
console.log((256n >> 3n).toString());
"#
        ),
        vec!["16", "32"]
    );
}

#[test]
fn bigint_typeof_is_bigint() {
    assert_eq!(
        run_js(
            r#"
console.log(typeof 42n);
"#
        ),
        vec!["bigint"]
    );
}

#[test]
fn bigint_to_string_with_radix() {
    assert_eq!(
        run_js(
            r#"
console.log((255n).toString(16));
console.log((8n).toString(2));
"#
        ),
        vec!["ff", "1000"]
    );
}

#[test]
fn bigint_throws_on_mixed_arithmetic_with_number() {
    assert_eq!(
        run_js(
            r#"
let threw = false;
try { 1n + 1; } catch { threw = true; }
console.log(threw);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn bigint_explicit_conversion() {
    assert_eq!(
        run_js(
            r#"
const n = 42n;
console.log(Number(n));
console.log(String(n));
const fromNum = BigInt(100);
console.log(fromNum === 100n);
"#
        ),
        vec!["42", "42", "true"]
    );
}

#[test]
fn bigint_negative() {
    assert_eq!(
        run_js(
            r#"
const n = -100n;
console.log(n < 0n);
console.log((-n).toString());
"#
        ),
        vec!["true", "100"]
    );
}

#[test]
fn bigint_as_object_key() {
    assert_eq!(
        run_js(
            r#"
const m = new Map();
m.set(1n, "one");
m.set(2n, "two");
console.log(m.get(1n));
console.log(m.size);
"#
        ),
        vec!["one", "2"]
    );
}
