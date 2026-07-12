use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Lambdas, delegates, AddressOf, closures
// ═══════════════════════════════════════════════════════════

#[test]
fn function_lambda_single_line() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim double As Func(Of Integer, Integer) = Function(x) x * 2
        Console.WriteLine(double(5))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn sub_lambda() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim greet As Action = Sub() Console.WriteLine("hello")
        greet()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn lambda_closure() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim counter As Integer = 0
        Dim inc As Action = Sub()
            counter = counter + 1
        End Sub
        inc()
        inc()
        inc()
        Console.WriteLine(counter)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn lambda_as_argument() {
    let out = run_vb(
        r#"
Module M
    Sub Apply(fn As Func(Of Integer, Integer), value As Integer)
        Console.WriteLine(fn(value))
    End Sub
    Sub Main()
        Apply(Function(x) x * x, 5)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["25"]);
}

#[test]
fn addressof_basic() {
    let out = run_vb(
        r#"
Module M
    Function Square(x As Integer) As Integer
        Return x * x
    End Function
    Sub Apply(fn As Func(Of Integer, Integer), value As Integer)
        Console.WriteLine(fn(value))
    End Sub
    Sub Main()
        Apply(AddressOf Square, 7)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["49"]);
}

#[test]
fn multiline_function_lambda() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim factorial As Func(Of Integer, Integer) = Function(n)
            If n <= 1 Then Return 1
            Return n * factorial(n - 1)
        End Function
        Console.WriteLine(factorial(5))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["120"]);
}

#[test]
fn lambda_returning_value() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim add As Func(Of Integer, Integer, Integer) = Function(a, b) a + b
        Console.WriteLine(add(3, 4))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["7"]);
}
