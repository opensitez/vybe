use super::helpers::run_python;

// decimal — Decimal, Context, getcontext, setcontext, localcontext, quantize, normalize, is_nan, is_infinite, rounding modes

#[test]
fn test_decimal_precision_context_setting() {
    let out = run_python(r#"
import decimal
ctx = decimal.getcontext()
ctx.prec = 6
d = decimal.Decimal(1) / decimal.Decimal(7)
print(d)
"#);
    assert_eq!(out, vec!["0.142857"]);
}

#[test]
fn test_decimal_quantize_rounding_half_up() {
    let out = run_python(r#"
import decimal
d = decimal.Decimal("2.675")
q = d.quantize(decimal.Decimal("0.01"), rounding=decimal.ROUND_HALF_UP)
print(q)
"#);
    assert_eq!(out, vec!["2.68"]);
}

#[test]
fn test_decimal_quantize_rounding_half_even_bankers() {
    let out = run_python(r#"
import decimal
d1 = decimal.Decimal("2.5")
d2 = decimal.Decimal("3.5")
q1 = d1.quantize(decimal.Decimal("1"), rounding=decimal.ROUND_HALF_EVEN)
q2 = d2.quantize(decimal.Decimal("1"), rounding=decimal.ROUND_HALF_EVEN)
print(q1, q2)
"#);
    assert_eq!(out, vec!["2 4"]);
}

#[test]
fn test_decimal_localcontext_scoped_precision() {
    let out = run_python(r#"
import decimal
decimal.getcontext().prec = 28
with decimal.localcontext() as ctx:
    ctx.prec = 4
    d1 = decimal.Decimal(1) / decimal.Decimal(3)
    print(d1)

d2 = decimal.Decimal(1) / decimal.Decimal(3)
print(len(str(d2)) > 10)
"#);
    assert_eq!(out, vec!["0.3333", "True"]);
}

#[test]
fn test_decimal_normalize_strips_trailing_zeros() {
    let out = run_python(r#"
import decimal
d = decimal.Decimal("1.5000")
norm = d.normalize()
print(norm)
"#);
    assert_eq!(out, vec!["1.5"]);
}

#[test]
fn test_decimal_is_nan_and_is_infinite() {
    let out = run_python(r#"
import decimal
d_nan = decimal.Decimal("NaN")
d_inf = decimal.Decimal("Infinity")
d_num = decimal.Decimal("123.45")
print(d_nan.is_nan())
print(d_inf.is_infinite())
print(d_num.is_finite())
"#);
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_decimal_as_tuple_representation() {
    let out = run_python(r#"
import decimal
d = decimal.Decimal("-12.34")
t = d.as_tuple()
print(t.sign)
print(t.digits)
print(t.exponent)
"#);
    assert_eq!(out, vec!["1", "(1, 2, 3, 4)", "-2"]);
}

#[test]
fn test_decimal_as_integer_ratio() {
    let out = run_python(r#"
import decimal
d = decimal.Decimal("0.75")
print(d.as_integer_ratio())
"#);
    assert_eq!(out, vec!["(3, 4)"]);
}

#[test]
fn test_decimal_sqrt_arbitrary_precision() {
    let out = run_python(r#"
import decimal
with decimal.localcontext() as ctx:
    ctx.prec = 10
    d = decimal.Decimal(2).sqrt()
    print(d)
"#);
    assert_eq!(out, vec!["1.414213562"]);
}

#[test]
fn test_decimal_ln_and_exp_functions() {
    let out = run_python(r#"
import decimal
with decimal.localcontext() as ctx:
    ctx.prec = 8
    e = decimal.Decimal(1).exp()
    print(str(e).startswith("2.71828"))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_decimal_exact_financial_addition() {
    let out = run_python(r#"
import decimal
val1 = decimal.Decimal("0.1")
val2 = decimal.Decimal("0.2")
print(val1 + val2 == decimal.Decimal("0.3"))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_decimal_comparison_with_floats_raises_typeerror_or_handles() {
    let out = run_python(r#"
import decimal
d = decimal.Decimal("1.5")
print(d == 1.5)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_decimal_to_integral_value_rounding() {
    let out = run_python(r#"
import decimal
d = decimal.Decimal("3.7")
print(d.to_integral_value(rounding=decimal.ROUND_FLOOR))
print(d.to_integral_value(rounding=decimal.ROUND_CEILING))
"#);
    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn test_decimal_remainder_and_remainder_near() {
    let out = run_python(r#"
import decimal
d1 = decimal.Decimal("10")
d2 = decimal.Decimal("3")
print(d1 % d2)
print(d1.remainder_near(d2))
"#);
    assert_eq!(out, vec!["1", "1"]);
}

#[test]
fn test_decimal_fma_fused_multiply_add() {
    let out = run_python(r#"
import decimal
d1 = decimal.Decimal("2")
d2 = decimal.Decimal("3")
d3 = decimal.Decimal("4")
res = d1.fma(d2, d3)  # 2 * 3 + 4
print(res)
"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_decimal_division_by_zero_signal() {
    let out = run_python(r#"
import decimal
try:
    with decimal.localcontext() as ctx:
        ctx.traps[decimal.DivisionByZero] = True
        decimal.Decimal(1) / decimal.Decimal(0)
except decimal.DivisionByZero:
    print("DivisionByZero")
"#);
    assert_eq!(out, vec!["DivisionByZero"]);
}

#[test]
fn test_decimal_create_decimal_from_float() {
    let out = run_python(r#"
import decimal
d = decimal.Decimal.from_float(0.5)
print(d)
"#);
    assert_eq!(out, vec!["0.5"]);
}

#[test]
fn test_decimal_max_and_min_methods() {
    let out = run_python(r#"
import decimal
d1 = decimal.Decimal("5.5")
d2 = decimal.Decimal("10.2")
print(d1.max(d2))
print(d1.min(d2))
"#);
    assert_eq!(out, vec!["10.2", "5.5"]);
}

#[test]
fn test_decimal_sign_and_copy_sign() {
    let out = run_python(r#"
import decimal
d1 = decimal.Decimal("-5")
d2 = decimal.Decimal("10")
print(d1.copy_sign(d2))
"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn test_decimal_adjusted_exponent_inspection() {
    let out = run_python(r#"
import decimal
d = decimal.Decimal("1.23e4")
print(d.adjusted())
"#);
    assert_eq!(out, vec!["4"]);
}
