use super::helpers::run_vb;

#[test]
fn floating_point_matrix_rounding_and_precision_contracts() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim values() As Double = {3.5, -3.5, 2.1, -2.1, 0.0, 1.499, 2.5}

        Console.WriteLine(Math.Round(3.5, 0))
        Console.WriteLine(Math.Round(2.5, 0))
        Console.WriteLine(Math.Ceiling(-3.5))
        Console.WriteLine(Math.Floor(2.5))
        Console.WriteLine(Math.Truncate(values(0)))
        Console.WriteLine(Math.Truncate(values(1)))
        Console.WriteLine(Math.Sign(values(2)))
        Console.WriteLine(Math.Sign(values(5)))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["4", "2", "-3", "2", "3", "-3", "1", "1"]);
}

#[test]
fn floating_point_matrix_exponential_logs_and_trig() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Console.WriteLine(Math.Round(Math.Log(Math.Exp(2)), 6))
        Console.WriteLine(Math.Round(Math.Sqrt(81), 6))
        Console.WriteLine(Math.Round(Math.Sin(0), 6))
        Console.WriteLine(Math.Round(Math.Cos(Math.PI), 6))
        Console.WriteLine(Math.Round(Math.Tan(Math.PI / 4), 6))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "9", "0", "-1", "1"]);
}

#[test]
fn floating_point_matrix_special_values_and_comparisons() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim n As Double = Double.NaN
        Dim p As Double = Double.PositiveInfinity
        Dim z As Double = Double.NegativeInfinity

        Console.WriteLine(Double.IsNaN(n))
        Console.WriteLine(Double.IsInfinity(p))
        Console.WriteLine(Double.IsInfinity(z))
        Console.WriteLine(n = n)
        Console.WriteLine(p > z)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "True", "False", "True"]);
}

#[test]
fn floating_point_matrix_epsilon_tolerance() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim left As Double = 0.1 + 0.2
        Dim expected As Double = 0.3

        Dim isClose As Boolean = Math.Abs(left - expected) < 1E-12
        Console.WriteLine(left = 0.3)
        Console.WriteLine(isClose)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn floating_point_matrix_integer_conversion_roundtrip() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim values() As Double = {1.2, -1.8, 9.99, 0.0}
        Dim allGood As Boolean = True

        For Each value In values
            Dim rounded As Double = Math.Round(value)
            Dim back As Double = CDbl(CInt(rounded))
            If back <> CInt(rounded) Then
                allGood = False
            End If
        Next

        Console.WriteLine(allGood)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}
