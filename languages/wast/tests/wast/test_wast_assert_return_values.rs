//! WAST script `assert_return` — actually EXECUTES the invoke and validates the
//! computed result against the expected value. Unlike the compile-only checks,
//! a wrong result (e.g. integer arithmetic that didn't wrap) fails the test.
use super::helpers::run_wast_asserts;

fn ok(src: &str) {
    run_wast_asserts(src).unwrap_or_else(|e| panic!("assertion failed: {e}\n{src}"));
}
fn must_fail(src: &str) {
    assert!(
        run_wast_asserts(src).is_err(),
        "expected a wrong result to be caught, but it passed:\n{src}"
    );
}

// ── the validator itself catches wrong results ───────────────────────────────
#[test]
fn wrong_i32_result_is_caught() {
    must_fail(
        r#"(module (func (export "f") (result i32) i32.const 41))
           (assert_return (invoke "f") (i32.const 42))"#,
    );
}
#[test]
fn correct_i32_result_passes() {
    ok(r#"(module (func (export "f") (result i32) i32.const 42))
          (assert_return (invoke "f") (i32.const 42))"#);
}

// ── arithmetic results verified through the assertion layer ───────────────────
#[test]
fn assert_i32_add() {
    ok(r#"(module (func (export "add") (param i32 i32) (result i32)
            local.get 0 local.get 1 i32.add))
          (assert_return (invoke "add" (i32.const 20) (i32.const 22)) (i32.const 42))"#);
}
#[test]
fn assert_i32_add_overflow_wraps() {
    ok(r#"(module (func (export "add") (param i32 i32) (result i32)
            local.get 0 local.get 1 i32.add))
          (assert_return (invoke "add" (i32.const 2147483647) (i32.const 1))
            (i32.const -2147483648))"#);
}
#[test]
fn assert_i32_mul_wraps() {
    ok(r#"(module (func (export "mul") (param i32 i32) (result i32)
            local.get 0 local.get 1 i32.mul))
          (assert_return (invoke "mul" (i32.const 65536) (i32.const 65536)) (i32.const 0))"#);
}
#[test]
fn assert_i32_div_u_unsigned() {
    ok(r#"(module (func (export "d") (param i32 i32) (result i32)
            local.get 0 local.get 1 i32.div_u))
          (assert_return (invoke "d" (i32.const -2) (i32.const 2)) (i32.const 2147483647))"#);
}
#[test]
fn assert_i64_arithmetic_exact() {
    ok(r#"(module (func (export "m") (param i64 i64) (result i64)
            local.get 0 local.get 1 i64.mul))
          (assert_return (invoke "m" (i64.const 1000000000) (i64.const 1000000000))
            (i64.const 1000000000000000000))"#);
}
#[test]
fn assert_i32_lt_u_unsigned_compare() {
    ok(r#"(module (func (export "lt") (param i32 i32) (result i32)
            local.get 0 local.get 1 i32.lt_u))
          (assert_return (invoke "lt" (i32.const -1) (i32.const 1)) (i32.const 0))"#);
}
#[test]
fn assert_i32_shr_s_sign_extends() {
    ok(r#"(module (func (export "s") (param i32 i32) (result i32)
            local.get 0 local.get 1 i32.shr_s))
          (assert_return (invoke "s" (i32.const -8) (i32.const 1)) (i32.const -4))"#);
}
#[test]
fn assert_f64_division() {
    ok(r#"(module (func (export "d") (param f64 f64) (result f64)
            local.get 0 local.get 1 f64.div))
          (assert_return (invoke "d" (f64.const 9.0) (f64.const 2.0)) (f64.const 4.5))"#);
}
#[test]
fn assert_i32_clz() {
    ok(
        r#"(module (func (export "c") (param i32) (result i32) local.get 0 i32.clz))
          (assert_return (invoke "c" (i32.const 1)) (i32.const 31))"#,
    );
}
#[test]
fn assert_i32_popcnt() {
    ok(
        r#"(module (func (export "p") (param i32) (result i32) local.get 0 i32.popcnt))
          (assert_return (invoke "p" (i32.const 255)) (i32.const 8))"#,
    );
}
#[test]
fn assert_multiple_invocations_same_module() {
    ok(
        r#"(module (func (export "sq") (param i32) (result i32) local.get 0 local.get 0 i32.mul))
          (assert_return (invoke "sq" (i32.const 3)) (i32.const 9))
          (assert_return (invoke "sq" (i32.const 5)) (i32.const 25))
          (assert_return (invoke "sq" (i32.const 12)) (i32.const 144))"#,
    );
}
#[test]
fn wrong_wrapping_result_is_caught() {
    // If i32.add did NOT wrap (returned 2147483648), this would be caught.
    must_fail(
        r#"(module (func (export "add") (param i32 i32) (result i32)
             local.get 0 local.get 1 i32.add))
           (assert_return (invoke "add" (i32.const 2147483647) (i32.const 1))
             (i32.const 2147483648))"#,
    );
}
