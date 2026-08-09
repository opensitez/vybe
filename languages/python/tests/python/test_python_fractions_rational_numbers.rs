use super::helpers::run_python;

// fractions — Fraction, numerator, denominator, limit_denominator, as_integer_ratio, arithmetic, math.gcd, string initialization, float conversion

#[test]
fn test_fractions_creation_and_simplification() {
    let out = run_python(
        r#"
from fractions import Fraction
f1 = Fraction(6, 8)
print(f1.numerator)
print(f1.denominator)
print(str(f1))
"#,
    );
    assert_eq!(out, vec!["3", "4", "3/4"]);
}

#[test]
fn test_fractions_string_parsing() {
    let out = run_python(
        r#"
from fractions import Fraction
f1 = Fraction("-3/7")
f2 = Fraction("2.5")
print(f1)
print(f2)
"#,
    );
    assert_eq!(out, vec!["-3/7", "5/2"]);
}

#[test]
fn test_fractions_limit_denominator_pi_approximation() {
    let out = run_python(
        r#"
from fractions import Fraction
import math

f_pi = Fraction(math.pi)
appr = f_pi.limit_denominator(100)
print(appr)
"#,
    );
    assert_eq!(out, vec!["22/7"]);
}

#[test]
fn test_fractions_as_integer_ratio_tuple() {
    let out = run_python(
        r#"
from fractions import Fraction
f = Fraction(11, 13)
print(f.as_integer_ratio())
"#,
    );
    assert_eq!(out, vec!["(11, 13)"]);
}

#[test]
fn test_fractions_arithmetic_ops() {
    let out = run_python(
        r#"
from fractions import Fraction
f1 = Fraction(1, 3)
f2 = Fraction(1, 6)
print(f1 + f2)
print(f1 - f2)
print(f1 * f2)
print(f1 / f2)
"#,
    );
    assert_eq!(out, vec!["1/2", "1/6", "1/18", "2"]);
}

#[test]
fn test_fractions_pow_exponentiation() {
    let out = run_python(
        r#"
from fractions import Fraction
f = Fraction(2, 3)
print(f ** 3)
print(f ** -2)
"#,
    );
    assert_eq!(out, vec!["8/27", "9/4"]);
}

#[test]
fn test_fractions_floor_ceil_round_trunc() {
    let out = run_python(
        r#"
from fractions import Fraction
import math

f = Fraction(7, 3)  # 2.333...
print(math.floor(f))
print(math.ceil(f))
print(round(f))
print(math.trunc(f))
"#,
    );
    assert_eq!(out, vec!["2", "3", "2", "2"]);
}

#[test]
fn test_fractions_zero_denominator_raises_zero_division() {
    let out = run_python(
        r#"
from fractions import Fraction
try:
    Fraction(1, 0)
except ZeroDivisionError:
    print("ZeroDivisionError")
"#,
    );
    assert_eq!(out, vec!["ZeroDivisionError"]);
}

#[test]
fn test_fractions_float_conversion_roundtrip() {
    let out = run_python(
        r#"
from fractions import Fraction
f = Fraction.from_float(0.125)
print(f)
print(float(f))
"#,
    );
    assert_eq!(out, vec!["1/8", "0.125"]);
}

#[test]
fn test_fractions_decimal_conversion() {
    let out = run_python(
        r#"
from fractions import Fraction
from decimal import Decimal

f = Fraction.from_decimal(Decimal("0.75"))
print(f)
"#,
    );
    assert_eq!(out, vec!["3/4"]);
}

#[test]
fn test_fractions_comparison_ops() {
    let out = run_python(
        r#"
from fractions import Fraction
f1 = Fraction(1, 2)
f2 = Fraction(2, 4)
f3 = Fraction(3, 4)
print(f1 == f2)
print(f1 < f3)
print(f3 > f2)
"#,
    );
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_fractions_hashability_set_dict_key() {
    let out = run_python(
        r#"
from fractions import Fraction
f1 = Fraction(1, 2)
f2 = Fraction(2, 4)
s = {f1, f2}
print(len(s))
d = {f1: "half"}
print(d[0.5])
"#,
    );
    assert_eq!(out, vec!["1", "half"]);
}

#[test]
fn test_fractions_mixed_number_with_integers() {
    let out = run_python(
        r#"
from fractions import Fraction
f = Fraction(3, 4)
print(f + 2)
print(5 - f)
print(f * 4)
"#,
    );
    assert_eq!(out, vec!["11/4", "17/4", "3"]);
}

#[test]
fn test_fractions_negative_signs_canonicalization() {
    let out = run_python(
        r#"
from fractions import Fraction
f1 = Fraction(3, -4)
f2 = Fraction(-3, -4)
print(f1)
print(f2)
"#,
    );
    assert_eq!(out, vec!["-3/4", "3/4"]);
}

#[test]
fn test_fractions_repr_formatting() {
    let out = run_python(
        r#"
from fractions import Fraction
f = Fraction(5, 6)
print(repr(f))
"#,
    );
    assert_eq!(out, vec!["Fraction(5, 6)"]);
}

#[test]
fn test_fractions_copy_and_pickle() {
    let out = run_python(
        r#"
import pickle
from fractions import Fraction

f1 = Fraction(7, 9)
data = pickle.dumps(f1)
f2 = pickle.loads(data)
print(f1 == f2)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_fractions_abs_and_neg() {
    let out = run_python(
        r#"
from fractions import Fraction
f = Fraction(-5, 8)
print(abs(f))
print(-f)
"#,
    );
    assert_eq!(out, vec!["5/8", "5/8"]);
}

#[test]
fn test_fractions_limit_denominator_max_denominator() {
    let out = run_python(
        r#"
from fractions import Fraction
f = Fraction("3.141592653589793")
lim = f.limit_denominator(10)
print(lim)
"#,
    );
    assert_eq!(out, vec!["22/7"]);
}

#[test]
fn test_fractions_int_conversion_truncation() {
    let out = run_python(
        r#"
from fractions import Fraction
f = Fraction(19, 5)  # 3.8
print(int(f))
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_fractions_gcd_helper_interoperability() {
    let out = run_python(
        r#"
from fractions import Fraction
import math

f1 = Fraction(12, 18)
print(math.gcd(f1.numerator, f1.denominator))
"#,
    );
    assert_eq!(out, vec!["1"]);
}
