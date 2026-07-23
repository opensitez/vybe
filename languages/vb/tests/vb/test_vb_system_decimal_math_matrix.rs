use super::helpers::run_vb;

#[test]
fn decimal_math_rounding_and_sign() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim a As Decimal = CDec(10)
        Dim b As Decimal = CDec(3)

        Console.WriteLine(Math.Round(a / b, 2))
        Console.WriteLine(Math.Sign(-a))
        Console.WriteLine(Math.Sign(0D))
        Console.WriteLine(Math.Sign(b))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3.33", "-1", "0", "1"]);
}

#[test]
fn decimal_math_add_subtract_and_remainder() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim a As Decimal = CDec(12.75)
        Dim b As Decimal = CDec(0.75)

        Console.WriteLine(a + b)
        Console.WriteLine(a - b)
        Console.WriteLine(a Mod b)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["13.5", "12", "0"]);
}

#[test]
fn decimal_math_comparisons_and_abs() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim a As Decimal = CDec(-12.34)

        Console.WriteLine(Decimal.Compare(a, Decimal.Zero) < 0)
        Console.WriteLine(Decimal.MaxValue > a)
        Console.WriteLine(Decimal.MinValue < a)
        Console.WriteLine(Decimal.Round(Math.Abs(a), 1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "True", "12.3"]);
}

#[test]
fn decimal_math_pow_log_not_available_use_squares() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim x As Decimal = CDec(2)
        Dim squared As Decimal = x * x
        Dim cubed As Decimal = x * x * x

        Console.WriteLine(squared)
        Console.WriteLine(cubed)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["4", "8"]);
}

#[test]
fn decimal_math_mixed_operations_cast_to_decimal() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim x As Decimal = CDec(1) / 3D
        Dim y As Decimal = CDec(2) * 3D

        Console.WriteLine(x = x)
        Console.WriteLine(y)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "6"]);
}
