// Python numeric tower — int, float, complex, Decimal, Fraction interop
use super::helpers::run_python;

#[test]
fn test_int_float_promotion() {
    let script = r#"
print(type(1 + 1.0).__name__)
print(type(10 / 3).__name__)
print(type(10 // 3).__name__)
"#;
    assert_eq!(run_python(script), vec!["float", "float", "int"]);
}

#[test]
fn test_complex_arithmetic() {
    let script = r#"
z1 = 3 + 4j
z2 = 1 - 2j
s = z1 + z2
print(s.real, s.imag)
p = z1 * z2
print(p.real, p.imag)
"#;
    assert_eq!(run_python(script), vec!["4.0 2.0", "11.0 -2.0"]);
}

#[test]
fn test_decimal_precision() {
    let script = r#"
from decimal import Decimal, getcontext
getcontext().prec = 5
a = Decimal('1') / Decimal('3')
print(str(a))
"#;
    assert_eq!(run_python(script), vec!["0.33333"]);
}

#[test]
fn test_fraction_exact_arithmetic() {
    let script = r#"
from fractions import Fraction
a = Fraction(1, 3)
b = Fraction(1, 6)
print(a + b)
print(a * b)
"#;
    assert_eq!(run_python(script), vec!["1/2", "1/18"]);
}

#[test]
fn test_int_division_modes() {
    let script = r#"
print(7 / 2)    # true div
print(7 // 2)   # floor div
print(-7 // 2)  # floor toward neg infinity
print(7 % 3)    # modulo
"#;
    assert_eq!(run_python(script), vec!["3.5", "3", "-4", "1"]);
}

#[test]
fn test_float_special_values() {
    let script = r#"
import math
inf = float('inf')
nan = float('nan')
print(math.isinf(inf))
print(math.isnan(nan))
print(inf > 1e308)
print(nan == nan)
"#;
    assert_eq!(run_python(script), vec!["True", "True", "True", "False"]);
}

#[test]
fn test_round_behavior() {
    let script = r#"
print(round(2.5))   # banker's rounding → 2
print(round(3.5))   # banker's rounding → 4
print(round(2.675, 2))  # float repr issue
print(round(-0.5))
"#;
    assert_eq!(run_python(script), vec!["2", "4", "2.67", "0"]);
}

#[test]
fn test_int_power_large() {
    let script = r#"
print(2 ** 100 > 10 ** 29)
print(type(2 ** 100).__name__)
"#;
    assert_eq!(run_python(script), vec!["True", "int"]);
}
