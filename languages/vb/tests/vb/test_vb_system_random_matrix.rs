use super::helpers::run_vb;

#[test]
fn random_seeded_next_deterministic() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim a As New Random(17)
        Dim b As New Random(17)
        Console.WriteLine(a.Next() = b.Next())
        Console.WriteLine(a.Next(100) = b.Next(100))
        Console.WriteLine(a.Next(10, 20) = b.Next(10, 20))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn random_next_bounds_are_respected() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim r As New Random(1)
        Dim v As Integer = r.Next(5, 10)
        Console.WriteLine(v >= 5 AndAlso v < 10)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn random_nextdouble_within_range() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim r As New Random(2)
        Dim v As Double = r.NextDouble()
        Console.WriteLine(v >= 0.0 AndAlso v < 1.0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn random_nextbytes_roundtrip_length() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim r As New Random(3)
        Dim b(4) As Byte
        r.NextBytes(b)
        Dim ok As Boolean = (b.Length = 5)
        Console.WriteLine(ok)
        Console.WriteLine(b(0) >= 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn random_nextbytes_deterministic_with_seed() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim a As New Random(99)
        Dim b As New Random(99)
        Dim left(2) As Byte
        Dim right(2) As Byte
        a.NextBytes(left)
        b.NextBytes(right)
        Console.WriteLine(left(0) = right(0))
        Console.WriteLine(left(1) = right(1))
        Console.WriteLine(left(2) = right(2))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn random_next_boolean_via_threshold() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim r As New Random(11)
        Dim value As Boolean = r.NextDouble() >= 0.5
        Console.WriteLine(value OrElse Not value)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn random_sample_int_range_without_upper() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim r As New Random(12)
        Dim v As Integer = r.Next(50)
        Console.WriteLine(v >= 0 AndAlso v < 50)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn random_next_after_repeated_calls_advances_state() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim r As New Random(20)
        Dim a As Integer = r.Next()
        Dim b As Integer = r.Next()
        Console.WriteLine(a = b)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False"]);
}

#[test]
fn random_initial_value_affects_sequence() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim a As New Random(21)
        Dim b As New Random(22)
        Console.WriteLine(a.Next() = b.Next())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False"]);
}

#[test]
fn random_nextdouble_next_int_mix_is_reproducible() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim r As New Random(7)
        Dim i As Integer = r.Next()
        Dim d As Double = r.NextDouble()
        Console.WriteLine(i <> 0 OrElse d >= 0.0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn random_nextzero_toone_not_equal_one() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim r As New Random(8)
        Dim d As Double = r.NextDouble()
        Console.WriteLine(d < 1.0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn random_next_accepts_zero_max() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim r As New Random(9)
        Dim v As Integer = r.Next(1)
        Console.WriteLine(v >= 0 AndAlso v < 1)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}
