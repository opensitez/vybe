//! Numeric precision: `decimal` exactness, `float` vs `double` precision, `BigInteger`.
use super::helpers::run_csharp;

#[test]
fn decimal_avoids_float_rounding_error() {
    assert_eq!(
        run_csharp(r#"decimal a=0.1m, b=0.2m;
Console.WriteLine(a+b==0.3m);"#),
        &["True"]
    );
}

#[test]
fn double_has_floating_point_rounding() {
    assert_eq!(
        run_csharp(r#"double a=0.1, b=0.2;
Console.WriteLine(a+b==0.3);"#),
        &["False"]
    );
}

#[test]
fn decimal_preserves_trailing_zeros_in_precision() {
    assert_eq!(
        run_csharp(r#"decimal d=1.50m;
Console.WriteLine(d.ToString(System.Globalization.CultureInfo.InvariantCulture));"#),
        &["1.50"]
    );
}

#[test]
fn float_is_32_bit_and_less_precise_than_double() {
    assert_eq!(
        run_csharp(r#"float f=1.0f/3.0f;
double d=1.0/3.0;
Console.WriteLine(f==(float)d);"#),
        &["True"]
    );
}

#[test]
fn big_integer_can_hold_arbitrarily_large_values() {
    assert_eq!(
        run_csharp(r#"var n=System.Numerics.BigInteger.Pow(10,30);
Console.WriteLine(n.ToString().StartsWith("1"));"#),
        &["True"]
    );
}

#[test]
fn big_integer_arithmetic_exact_for_large_factorial() {
    assert_eq!(
        run_csharp(r#"System.Numerics.BigInteger f=1;
for(int i=1;i<=20;i++) f*=i;
Console.WriteLine(f.ToString());"#),
        &["2432902008176640000"]
    );
}
