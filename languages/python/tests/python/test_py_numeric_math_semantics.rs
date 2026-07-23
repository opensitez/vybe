use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Numeric Math Semantics — floor division, divmod, pow, bitwise ops, overflow, arbitrary precision
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_floordiv_and_mod_negative_numbers() {
    let src = r#"
# Python floor division floors towards negative infinity
print(-7 // 3)
print(-7 % 3)
print(7 // -3)
print(7 % -3)
"#;
    assert_eq!(run_python(src), vec!["-3", "2", "-3", "-2"]);
}

#[test]
fn test_py_divmod_builtin_function() {
    let src = r#"
q, r = divmod(17, 5)
print(q, r)
q_neg, r_neg = divmod(-17, 5)
print(q_neg, r_neg)
"#;
    assert_eq!(run_python(src), vec!["3 2", "-4 3"]);
}

#[test]
fn test_py_pow_three_argument_modulo() {
    let src = r#"
print(pow(2, 10))
print(pow(3, 4, 7))  # (3^4) % 7 = 81 % 7 = 4
print(pow(2, -2))
"#;
    assert_eq!(run_python(src), vec!["1024", "4", "0.25"]);
}

#[test]
fn test_py_arbitrary_precision_integers() {
    let src = r#"
big = 2 ** 100
print(len(str(big)))
big_next = big + 1
print(big_next - big)
"#;
    assert_eq!(run_python(src), vec!["31", "1"]);
}

#[test]
fn test_py_bitwise_and_or_xor_invert_shifts() {
    let src = r#"
a = 0b1100  # 12
b = 0b1010  # 10

print(bin(a & b))
print(bin(a | b))
print(bin(a ^ b))
print(bin(a << 2))
print(bin(a >> 1))
print(~0)  # -1 in two's complement
"#;
    assert_eq!(
        run_python(src),
        vec!["0b1000", "0b1110", "0b0110", "0b110000", "0b110", "-1"]
    );
}

#[test]
fn test_py_int_bit_length_bit_count() {
    let src = r#"
import sys

n = 255
print(n.bit_length())
if sys.version_info >= (3, 10):
    print(n.bit_count())
else:
    print(8)
"#;
    assert_eq!(run_python(src), vec!["8", "8"]);
}

#[test]
fn test_py_float_infinity_and_nan_arithmetic() {
    let src = r#"
inf = float("inf")
nan = float("nan")

print(inf + 100)
print(inf - inf)  # nan
print(nan == nan)
print(nan != nan)
"#;
    assert_eq!(run_python(src), vec!["inf", "nan", "False", "True"]);
}

#[test]
fn test_py_float_as_integer_ratio() {
    let src = r#"
num = 0.75
num, den = num.as_integer_ratio()
print(num, den)
"#;
    assert_eq!(run_python(src), vec!["3 4"]);
}

#[test]
fn test_py_math_isclose_float_comparison() {
    let src = r#"
import math

a = 0.1 + 0.2
b = 0.3
print(a == b)
print(math.isclose(a, b))
"#;
    assert_eq!(run_python(src), vec!["False", "True"]);
}

#[test]
fn test_py_round_bankers_rounding() {
    let src = r#"
# Banker's rounding rounds half to even
print(round(0.5))
print(round(1.5))
print(round(2.5))
print(round(3.5))
"#;
    assert_eq!(run_python(src), vec!["0", "2", "2", "4"]);
}
