use super::helpers::{compile_ok, run_ruby, run_ruby_one};

// ── Integer iteration ────────────────────────────────────────────────────────

#[test]
fn times_loop() {
    compile_ok("3.times { |i| puts i }\n");
}

#[test]
fn upto_loop() {
    compile_ok("1.upto(5) { |i| puts i }\n");
}

#[test]
fn downto_loop() {
    compile_ok("5.downto(1) { |i| puts i }\n");
}

// ── Base conversion ──────────────────────────────────────────────────────────

#[test]
fn to_s_binary() {
    compile_ok("x = 10.to_s(2)\n");
}

#[test]
fn to_s_octal() {
    compile_ok("x = 8.to_s(8)\n");
}

#[test]
fn to_s_hex() {
    compile_ok("x = 255.to_s(16)\n");
}

// ── Integer arithmetic helpers ───────────────────────────────────────────────

#[test]
fn divmod_call() {
    compile_ok("q, r = 17.divmod(5)\n");
}

#[test]
fn gcd_call() {
    compile_ok("x = 12.gcd(8)\n");
}

#[test]
fn lcm_call() {
    compile_ok("x = 4.lcm(6)\n");
}

// ── Float precision methods ──────────────────────────────────────────────────

#[test]
fn floor_with_precision() {
    compile_ok("x = 3.14159.floor(2)\n");
}

#[test]
fn ceil_with_precision() {
    compile_ok("x = 3.14159.ceil(2)\n");
}

#[test]
fn round_with_precision() {
    compile_ok("x = 3.14159.round(3)\n");
}

#[test]
fn round_half_up_mode() {
    compile_ok("x = 2.5.round(half: :up)\n");
}

#[test]
fn truncate_with_precision() {
    compile_ok("x = 3.9999.truncate(2)\n");
}

// ── Clamp ────────────────────────────────────────────────────────────────────

#[test]
fn clamp_min_max() {
    compile_ok("x = 15.clamp(0, 10)\n");
}

#[test]
fn clamp_range() {
    compile_ok("x = 15.clamp(0..10)\n");
}

// ── Digit decomposition ──────────────────────────────────────────────────────

#[test]
fn digits_base10() {
    compile_ok("x = 123.digits\n");
}

#[test]
fn digits_base() {
    compile_ok("x = 10.digits(2)\n");
}

// ── Character conversion ─────────────────────────────────────────────────────

#[test]
fn chr_call() {
    compile_ok("x = 65.chr\n");
}

// ── Strict coercion functions ────────────────────────────────────────────────

#[test]
fn integer_coerce() {
    compile_ok("x = Integer('42')\n");
}

#[test]
fn float_coerce() {
    compile_ok("x = Float('3.14')\n");
}

// ── Rational and Complex (compile-only; no VM support required) ───────────────

#[test]
fn rational_literal() {
    compile_ok("x = Rational(3, 4)\n");
}

#[test]
fn complex_literal() {
    compile_ok("x = Complex(2, 3)\n");
}

// ── Float predicates ─────────────────────────────────────────────────────────

#[test]
fn float_infinite() {
    compile_ok("x = (1.0 / 0).infinite?\n");
}

#[test]
fn float_nan() {
    compile_ok("x = (0.0 / 0).nan?\n");
}

#[test]
fn float_finite() {
    compile_ok("x = 3.14.finite?\n");
}

// ── Negative float ceil/floor asymmetry ──────────────────────────────────────

#[test]
fn ceil_negative_float() {
    compile_ok("x = (-2.3).ceil\n");
}

#[test]
fn floor_negative_float() {
    compile_ok("x = (-2.3).floor\n");
}

// ── Complex abs2 (compile-only) ───────────────────────────────────────────────

#[test]
fn complex_abs2() {
    compile_ok("c = Complex(3, 4)\nx = c.abs2\n");
}

// ── Modular exponentiation ───────────────────────────────────────────────────

#[test]
fn pow_with_mod() {
    compile_ok("x = 2.pow(10, 1000)\n");
}

// ── Bit length ───────────────────────────────────────────────────────────────

#[test]
fn bit_length_call() {
    compile_ok("x = 255.bit_length\n");
}

// ── Runtime spot-checks ──────────────────────────────────────────────────────

#[test]
fn times_loop_runtime() {
    let out = run_ruby("3.times { |i| puts i }\n");
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn upto_loop_runtime() {
    let out = run_ruby("1.upto(3) { |i| puts i }\n");
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn downto_loop_runtime() {
    let out = run_ruby("3.downto(1) { |i| puts i }\n");
    assert_eq!(out, vec!["3", "2", "1"]);
}

#[test]
fn gcd_runtime() {
    assert_eq!(run_ruby_one("puts 12.gcd(8)\n"), "4");
}

#[test]
fn lcm_runtime() {
    assert_eq!(run_ruby_one("puts 4.lcm(6)\n"), "12");
}

#[test]
fn clamp_above_max_runtime() {
    assert_eq!(run_ruby_one("puts 15.clamp(0, 10)\n"), "10");
}

#[test]
fn clamp_within_range_runtime() {
    assert_eq!(run_ruby_one("puts 5.clamp(0, 10)\n"), "5");
}

#[test]
fn chr_runtime() {
    assert_eq!(run_ruby_one("puts 65.chr\n"), "A");
}
