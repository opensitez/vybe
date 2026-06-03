use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Operators — comparison, logical, arithmetic, bitwise
// ═══════════════════════════════════════════════════════════

#[test]
fn comparison_operators() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(1 < 2)
        Console.WriteLine(2 > 1)
        Console.WriteLine(1 <= 1)
        Console.WriteLine(1 >= 1)
        Console.WriteLine(1 = 1)
        Console.WriteLine(1 <> 2)
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        super::helpers::dotnet_expected_lines(&["true", "true", "true", "true", "true", "true"])
    );
}

#[test]
fn logical_and_or_not() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(True And True)
        Console.WriteLine(True And False)
        Console.WriteLine(False Or True)
        Console.WriteLine(False Or False)
        Console.WriteLine(Not True)
        Console.WriteLine(Not False)
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        super::helpers::dotnet_expected_lines(&["true", "false", "true", "false", "false", "true"])
    );
}

#[test]
fn andalso_orelse_short_circuit() {
    let out = run_vb(
        r#"
Module M
    Dim called As Boolean = False
    Function SideEffect() As Boolean
        called = True
        Return True
    End Function
    Sub Main()
        ' AndAlso short-circuits: SideEffect should NOT be called
        If False AndAlso SideEffect() Then
            Console.WriteLine("never")
        End If
        Console.WriteLine(called)
        ' OrElse short-circuits: SideEffect should NOT be called
        called = False
        If True OrElse SideEffect() Then
            Console.WriteLine("yes")
        End If
        Console.WriteLine(called)
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        super::helpers::dotnet_expected_lines(&["false", "yes", "false"])
    );
}

#[test]
fn arithmetic_basic() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(10 + 5)
        Console.WriteLine(10 - 5)
        Console.WriteLine(10 * 5)
        Console.WriteLine(10 / 4)
        Console.WriteLine(10 \ 3)
        Console.WriteLine(10 Mod 3)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["15", "5", "50", "2.5", "3", "1"]);
}

#[test]
fn exponentiation() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(2 ^ 10)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1024"]);
}

#[test]
fn compound_assignment() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim x As Integer = 10
        x += 5
        Console.WriteLine(x)
        x -= 3
        Console.WriteLine(x)
        x *= 2
        Console.WriteLine(x)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["15", "12", "24"]);
}

#[test]
fn string_concat_ampersand() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim result As String = "Hello" & " " & "World"
        Console.WriteLine(result)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Hello World"]);
}

#[test]
fn concat_assign() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim s As String = "Hello"
        s &= " World"
        Console.WriteLine(s)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Hello World"]);
}

#[test]
fn unary_minus() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim x As Integer = 5
        Console.WriteLine(-x)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["-5"]);
}

#[test]
fn parenthesized_expressions() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine((2 + 3) * 4)
        Console.WriteLine(2 + 3 * 4)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["20", "14"]);
}

#[test]
fn math_functions() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Abs(-42))
        Console.WriteLine(Math.Max(10, 20))
        Console.WriteLine(Math.Min(10, 20))
        Console.WriteLine(Math.Sqrt(25))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["42", "20", "10", "5"]);
}

#[test]
fn math_floor_ceiling() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Floor(3.7))
        Console.WriteLine(Math.Ceiling(3.2))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3", "4"]);
}
