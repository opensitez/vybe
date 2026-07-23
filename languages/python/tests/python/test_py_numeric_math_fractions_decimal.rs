use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Numeric Math, Fractions & Decimal — Decimal precision, rounding modes, Fraction exact math, math utilities
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_decimal_exact_arithmetic_precision() {
    let src = r#"
from decimal import Decimal, getcontext

getcontext().prec = 6
d1 = Decimal("1.234567")
d2 = Decimal("2.345678")
print(d1 + d2)
"#;
    assert_eq!(run_python(src), vec!["3.58025"]);
}

#[test]
fn test_py_decimal_rounding_modes_quantize() {
    let src = r#"
from decimal import Decimal, ROUND_HALF_UP, ROUND_HALF_EVEN, ROUND_FLOOR

d = Decimal("2.5")
print(d.quantize(Decimal("1"), rounding=ROUND_HALF_UP))
print(d.quantize(Decimal("1"), rounding=ROUND_HALF_EVEN))
print(d.quantize(Decimal("1"), rounding=ROUND_FLOOR))
"#;
    assert_eq!(run_python(src), vec!["3", "2", "2"]);
}

#[test]
fn test_py_fraction_exact_rational_arithmetic() {
    let src = r#"
from fractions import Fraction

f1 = Fraction(1, 3)
f2 = Fraction(1, 6)
sum_f = f1 + f2
print(sum_f.numerator, sum_f.denominator)
print(sum_f)
"#;
    assert_eq!(run_python(src), vec!["1 2", "1/2"]);
}

#[test]
fn test_py_fraction_conversion_from_string_float() {
    let src = r#"
from fractions import Fraction

f_str = Fraction("0.75")
f_float = Fraction(0.5)
print(f_str)
print(f_float)
"#;
    assert_eq!(run_python(src), vec!["3/4", "1/2"]);
}

#[test]
fn test_py_math_gcd_lcm_comb_perm() {
    let src = r#"
import math

print(math.gcd(24, 36))
print(math.lcm(4, 6))
print(math.comb(5, 2))
print(math.perm(5, 2))
"#;
    assert_eq!(run_python(src), vec!["12", "12", "10", "20"]);
}

#[test]
fn test_py_math_isclose_rel_abs_tol() {
    let src = r#"
import math

a = 100.0
b = 100.00001
print(math.isclose(a, b, rel_tol=1e-5))
print(math.isclose(a, b, rel_tol=1e-7))
"#;
    assert_eq!(run_python(src), vec!["True", "False"]);
}

#[test]
fn test_py_complex_number_polar_coordinates() {
    let src = r#"
import cmath

z = complex(3, 4)
r, theta = cmath.polar(z)
print(round(r, 2))
z_rect = cmath.rect(r, theta)
print(round(z_rect.real, 2), round(z_rect.imag, 2))
"#;
    assert_eq!(run_python(src), vec!["5.0", "3.0 4.0"]);
}

#[test]
fn test_py_math_hypot_3d_dist() {
    let src = r#"
import math

print(math.hypot(3, 4))
print(math.hypot(1, 2, 2))  # 3D
print(math.dist([0, 0], [3, 4]))
"#;
    assert_eq!(run_python(src), vec!["5.0", "3.0", "5.0"]);
}

#[test]
fn test_py_decimal_tuple_construction() {
    let src = r#"
from decimal import Decimal

# tuple: (sign, digits_tuple, exponent) -> 0 means positive, (1, 2, 5) is 125, exp -2 -> 1.25
d = Decimal((0, (1, 2, 5), -2))
print(d)
"#;
    assert_eq!(run_python(src), vec!["1.25"]);
}

#[test]
fn test_py_math_remainder_fmod() {
    let src = r#"
import math

print(math.fmod(7, 3))
print(math.remainder(7, 3))  # IEEE 754 remainder
"#;
    assert_eq!(run_python(src), vec!["1.0", "1.0"]);
}
