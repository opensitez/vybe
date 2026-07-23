use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: math + numbers — math module, decimal, fractions, statistics, complex numbers, int methods
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_math_basic_functions() {
    let src = r#"
import math

print(math.sqrt(16))
print(math.ceil(4.3))
print(math.floor(4.7))
print(math.trunc(4.9))
print(math.fabs(-5.5))
print(math.factorial(5))
"#;
    assert_eq!(run_python(src), vec!["4.0", "5", "4", "4", "5.5", "120"]);
}

#[test]
fn test_py_math_trigonometry() {
    let src = r#"
import math

print(round(math.sin(math.pi / 2), 5))
print(round(math.cos(0), 5))
print(round(math.tan(math.pi / 4), 5))
print(round(math.degrees(math.pi), 5))
print(round(math.radians(180), 5))
"#;
    assert_eq!(
        run_python(src),
        vec!["1.0", "1.0", "1.0", "180.0", "3.14159"]
    );
}

#[test]
fn test_py_math_logarithms_and_exp() {
    let src = r#"
import math

print(round(math.log(math.e), 5))
print(round(math.log(100, 10), 5))
print(round(math.log2(8), 5))
print(round(math.log10(1000), 5))
print(round(math.exp(1), 5))
"#;
    assert_eq!(run_python(src), vec!["1.0", "2.0", "3.0", "3.0", "2.71828"]);
}

#[test]
fn test_py_math_special_functions() {
    let src = r#"
import math

print(math.gcd(12, 8))
print(math.gcd(100, 75))
print(math.lcm(4, 6))
print(math.isfinite(float('inf')))
print(math.isinf(float('inf')))
print(math.isnan(float('nan')))
print(math.copysign(3, -1))
"#;
    assert_eq!(
        run_python(src),
        vec!["4", "25", "12", "False", "True", "True", "-3.0"]
    );
}

#[test]
fn test_py_math_hypot_and_dist() {
    let src = r#"
import math

print(round(math.hypot(3, 4), 5))
print(round(math.hypot(1, 1, 1), 5))  # 3D
print(round(math.dist([0, 0], [3, 4]), 5))
"#;
    assert_eq!(run_python(src), vec!["5.0", "1.73205", "5.0"]);
}

#[test]
fn test_py_decimal_precision_arithmetic() {
    let src = r#"
from decimal import Decimal, getcontext

getcontext().prec = 10
a = Decimal("1.1")
b = Decimal("2.2")
print(a + b)
print(Decimal("0.1") + Decimal("0.2"))
print(float(0.1) + float(0.2))  # floating point imprecision
"#;
    assert_eq!(run_python(src), vec!["3.3", "0.3", "0.30000000000000004"]);
}

#[test]
fn test_py_decimal_rounding_modes() {
    let src = r#"
from decimal import Decimal, ROUND_HALF_UP, ROUND_HALF_EVEN

d = Decimal("2.5")
print(d.quantize(Decimal("1"), rounding=ROUND_HALF_UP))
print(d.quantize(Decimal("1"), rounding=ROUND_HALF_EVEN))

d2 = Decimal("3.5")
print(d2.quantize(Decimal("1"), rounding=ROUND_HALF_EVEN))
"#;
    assert_eq!(run_python(src), vec!["3", "2", "4"]);
}

#[test]
fn test_py_fractions_exact_arithmetic() {
    let src = r#"
from fractions import Fraction

a = Fraction(1, 3)
b = Fraction(1, 6)
print(a + b)
print(a * b)
print(a - b)
print(Fraction("1.5"))
print(Fraction(0.25))
"#;
    assert_eq!(run_python(src), vec!["1/2", "1/18", "1/6", "3/2", "1/4"]);
}

#[test]
fn test_py_statistics_basic() {
    let src = r#"
import statistics

data = [1, 2, 3, 4, 5, 5, 6, 7, 8]
print(statistics.mean(data))
print(statistics.median(data))
print(statistics.mode(data))
print(round(statistics.stdev(data), 4))
print(round(statistics.variance(data), 4))
"#;
    assert_eq!(
        run_python(src),
        vec!["4.555555555555555", "5", "5", "2.1268", "4.5278"]
    );
}

#[test]
fn test_py_statistics_quantiles() {
    let src = r#"
import statistics

data = list(range(1, 101))
q = statistics.quantiles(data, n=4)
print(q[0])   # Q1
print(q[1])   # median
print(q[2])   # Q3
"#;
    assert_eq!(run_python(src), vec!["25.75", "50.5", "75.25"]);
}

#[test]
fn test_py_complex_number_operations() {
    let src = r#"
z1 = 3 + 4j
z2 = 1 - 2j
print(z1 + z2)
print(z1 * z2)
print(z1.real, z1.imag)
print(abs(z1))
print(z1.conjugate())
"#;
    assert_eq!(
        run_python(src),
        vec!["(4+2j)", "(11-2j)", "3.0 4.0", "5.0", "(3-4j)"]
    );
}

#[test]
fn test_py_int_bit_operations() {
    let src = r#"
n = 255
print(n.bit_length())
print(bin(n))

large = 2 ** 100
print(large.bit_length())
print(int.from_bytes(b'\xff', byteorder='big'))
print((256).to_bytes(2, byteorder='big'))
"#;
    assert_eq!(
        run_python(src),
        vec!["8", "0b11111111", "101", "255", "b'\\x01\\x00'"]
    );
}

#[test]
fn test_py_math_constants() {
    let src = r#"
import math

print(round(math.pi, 5))
print(round(math.e, 5))
print(round(math.tau, 5))
print(math.inf > 1e308)
print(math.nan != math.nan)  # NaN != NaN
"#;
    assert_eq!(
        run_python(src),
        vec!["3.14159", "2.71828", "6.28318", "True", "True"]
    );
}

#[test]
fn test_py_math_comb_perm() {
    let src = r#"
import math

print(math.comb(10, 3))    # 10 choose 3 = 120
print(math.perm(10, 3))    # 10 * 9 * 8 = 720
print(math.comb(5, 5))     # = 1
print(math.comb(5, 0))     # = 1
"#;
    assert_eq!(run_python(src), vec!["120", "720", "1", "1"]);
}
