//! Numeric operators: modulo, integer division, mixed types, widening.
use super::helpers::run_csharp;

#[test]
fn integer_division_truncates_toward_zero() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(7/2);
Console.WriteLine(-7/2);"#),
        &["3", "-3"]
    );
}

#[test]
fn modulo_sign_follows_dividend_not_divisor() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(7%3);
Console.WriteLine(-7%3);"#),
        &["1", "-1"]
    );
}

#[test]
fn integer_plus_double_widens_to_double() {
    assert_eq!(
        run_csharp(r#"int i=3; double d=1.5;
Console.WriteLine(i+d);"#),
        &["4.5"]
    );
}

#[test]
fn integer_division_by_zero_throws() {
    assert_eq!(
        run_csharp(r#"string r="";
try{int x=1/0;}
catch(System.DivideByZeroException){r="div0";}
Console.WriteLine(r);"#),
        &["div0"]
    );
}

#[test]
fn double_division_by_zero_produces_infinity() {
    assert_eq!(
        run_csharp(r#"double d=1.0/0.0;
Console.WriteLine(double.IsInfinity(d));"#),
        &["True"]
    );
}

#[test]
fn long_arithmetic_handles_large_values_exactly() {
    assert_eq!(
        run_csharp(r#"long a=3_000_000_000L; long b=a*2;
Console.WriteLine(b);"#),
        &["6000000000"]
    );
}

#[test]
fn postfix_increment_returns_old_value() {
    assert_eq!(
        run_csharp(r#"int x=5; int y=x++;
Console.WriteLine(y); Console.WriteLine(x);"#),
        &["5", "6"]
    );
}

#[test]
fn prefix_increment_returns_new_value() {
    assert_eq!(
        run_csharp(r#"int x=5; int y=++x;
Console.WriteLine(y); Console.WriteLine(x);"#),
        &["6", "6"]
    );
}
