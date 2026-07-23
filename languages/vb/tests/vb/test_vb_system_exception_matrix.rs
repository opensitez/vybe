use super::helpers::run_vb;

#[test]
fn exception_divide_by_zero_is_caught() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Try
            Dim zero As Integer = 0
            Console.WriteLine(1 \ zero)
        Catch ex As DivideByZeroException
            Console.WriteLine("zero")
        End Try
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["zero"]);
}

#[test]
fn exception_invalid_cast_is_reported_as_type() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Try
            Dim value As Object = "text"
            Dim asInt As Integer = CInt(value)
            Console.WriteLine("ok")
        Catch ex As InvalidCastException
            Console.WriteLine(ex.GetType().Name)
        End Try
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["InvalidCastException"]);
}

#[test]
fn exception_general_catch_always_runs_for_failure() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Try
            Dim arr(0) As Integer
            Console.WriteLine(arr(1))
        Catch ex As Exception
            Console.WriteLine(ex.GetType().Name)
        End Try
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["IndexOutOfRangeException"]);
}

#[test]
fn exception_finally_executes_once() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim flag As Integer = 0
        Try
            flag = 1
        Finally
            flag = 2
        End Try
        Console.WriteLine(flag)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2"]);
}

#[test]
fn exception_rethrow_preserves_type() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Try
            Try
                Throw New ArgumentException("inner")
            Catch ex As Exception
                Throw
            End Try
        Catch ex As Exception
            Console.WriteLine(ex.GetType().Name)
            Console.WriteLine(ex.Message.Contains("inner"))
        End Try
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["ArgumentException", "True"]);
}

#[test]
fn exception_when_filter_matches_type() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Try
            Throw New ArgumentOutOfRangeException("x")
        Catch ex As Exception When TypeOf ex Is ArgumentOutOfRangeException
            Console.WriteLine("filtered")
        End Try
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["filtered"]);
}

#[test]
fn exception_nested_try_finally_contract() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim depth As Integer = 0
        Try
            Try
                depth = 1
            Finally
                depth += 1
            End Try
        Catch ex As Exception
            depth = -1
        End Try
        Console.WriteLine(depth)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2"]);
}
