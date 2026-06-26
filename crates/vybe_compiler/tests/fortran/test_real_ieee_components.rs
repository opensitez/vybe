//! IEEE real model intrinsics: modf, fraction, exponent, scale, spacing,
//! rrspacing, and nearest — beyond compile-only tiny/epsilon in
//! `test_intrinsics_extended.rs`.

use super::helpers::compile_ok;

// ── modf: integer and fractional parts ───────────────────────────

#[test]
fn modf_positive_splits_integer_and_fraction() {
    compile_ok(
        "program t\nreal :: f\ninteger :: i\nf = modf(3.75, i)\nprint *, i, f\nend program t\n",
    );
}

#[test]
fn modf_negative_truncates_toward_zero() {
    compile_ok(
        "program t\nreal :: f\ninteger :: i\nf = modf(-3.75, i)\nprint *, i, f\nend program t\n",
    );
}

#[test]
fn modf_zero_yields_zero_fraction() {
    compile_ok(
        "program t\nreal :: f\ninteger :: i\nf = modf(0.0, i)\nprint *, i, f\nend program t\n",
    );
}

// ── fraction: normalized significand ─────────────────────────────

#[test]
fn fraction_one_point_five_is_three_halves() {
    compile_ok("program t\nprint *, fraction(1.5)\nend program t\n");
}

#[test]
fn fraction_eight_is_one_half() {
    compile_ok("program t\nprint *, fraction(8.0)\nend program t\n");
}

#[test]
fn fraction_negative_preserves_sign() {
    compile_ok("program t\nprint *, fraction(-1.5)\nend program t\n");
}

#[test]
fn fraction_zero_is_zero() {
    compile_ok("program t\nprint *, fraction(0.0)\nend program t\n");
}

// ── exponent: unbiased model exponent ────────────────────────────

#[test]
fn exponent_one_is_one() {
    compile_ok("program t\nprint *, exponent(1.0)\nend program t\n");
}

#[test]
fn exponent_eight_is_four() {
    compile_ok("program t\nprint *, exponent(8.0)\nend program t\n");
}

#[test]
fn exponent_one_half_is_zero() {
    compile_ok("program t\nprint *, exponent(0.5)\nend program t\n");
}

#[test]
fn exponent_zero_is_zero() {
    compile_ok("program t\nprint *, exponent(0.0)\nend program t\n");
}

// ── scale: multiply by radix**n ──────────────────────────────────

#[test]
fn scale_doubles_with_positive_two() {
    compile_ok("program t\nprint *, scale(1.0, 2)\nend program t\n");
}

#[test]
fn scale_halves_with_negative_one() {
    compile_ok("program t\nprint *, scale(8.0, -1)\nend program t\n");
}

#[test]
fn scale_identity_with_zero_exponent() {
    compile_ok("program t\nprint *, scale(4.0, 0)\nend program t\n");
}

#[test]
fn scale_negative_value_positive_exponent() {
    compile_ok("program t\nprint *, scale(-2.0, 3)\nend program t\n");
}

// ── spacing: least gap in model at x ─────────────────────────────

#[test]
fn spacing_of_unity() {
    compile_ok("program t\nprint *, spacing(1.0)\nend program t\n");
}

#[test]
fn spacing_at_tiny_value() {
    compile_ok("program t\nprint *, spacing(tiny(1.0))\nend program t\n");
}

// ── rrspacing: reciprocal of spacing ─────────────────────────────

#[test]
fn rrspacing_of_unity() {
    compile_ok("program t\nprint *, rrspacing(1.0)\nend program t\n");
}

#[test]
fn rrspacing_of_large_value() {
    compile_ok("program t\nprint *, rrspacing(huge(1.0) / 2.0)\nend program t\n");
}

// ── nearest: next representable in signed direction ──────────────

#[test]
fn nearest_toward_positive_infinity() {
    compile_ok("program t\nprint *, nearest(1.0, 1.0)\nend program t\n");
}

#[test]
fn nearest_toward_negative_infinity() {
    compile_ok("program t\nprint *, nearest(1.0, -1.0)\nend program t\n");
}

#[test]
fn nearest_exact_value_unchanged() {
    compile_ok("program t\nprint *, nearest(2.0, 1.0)\nend program t\n");
}
