use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Error handling — Try/Catch/Finally
// ═══════════════════════════════════════════════════════════

#[test]
fn try_catch_basic() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Try
            Throw New Exception("oops")
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["oops"]);
}

#[test]
fn try_catch_no_error() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Try
            Console.WriteLine("ok")
        Catch ex As Exception
            Console.WriteLine("error")
        End Try
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn try_finally() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Try
            Console.WriteLine("try")
        Finally
            Console.WriteLine("finally")
        End Try
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["try", "finally"]);
}

#[test]
fn try_catch_finally() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Try
            Throw New Exception("fail")
        Catch ex As Exception
            Console.WriteLine("caught: " & ex.Message)
        Finally
            Console.WriteLine("cleanup")
        End Try
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["caught: fail", "cleanup"]);
}

#[test]
fn throw_and_catch() {
    let out = run_vb(
        r#"
Module M
    Sub Risky()
        Throw New Exception("danger")
    End Sub
    Sub Main()
        Try
            Risky()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["danger"]);
}

#[test]
fn nested_try_catch() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Try
            Try
                Throw New Exception("inner")
            Catch ex As Exception
                Console.WriteLine("inner: " & ex.Message)
                Throw New Exception("rethrown")
            End Try
        Catch ex As Exception
            Console.WriteLine("outer: " & ex.Message)
        End Try
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["inner: inner", "outer: rethrown"]);
}

#[test]
fn finally_always_runs() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Try
            Console.WriteLine("before")
            Throw New Exception("err")
        Catch ex As Exception
            Console.WriteLine("caught")
        Finally
            Console.WriteLine("always")
        End Try
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["before", "caught", "always"]);
}

#[test]
fn exception_in_function() {
    let out = run_vb(
        r#"
Module M
    Function Divide(a As Integer, b As Integer) As Integer
        If b = 0 Then
            Throw New Exception("Division by zero")
        End If
        Return a \ b
    End Function
    Sub Main()
        Try
            Console.WriteLine(Divide(10, 2))
            Console.WriteLine(Divide(10, 0))
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["5", "Division by zero"]);
}
