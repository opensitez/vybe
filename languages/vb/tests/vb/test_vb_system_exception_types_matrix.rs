use super::helpers::run_vb;

#[test]
fn exception_argument_exception_exposes_parameter_name() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Try
            Throw New ArgumentException("missing value", "value")
        Catch ex As ArgumentException
            Console.WriteLine(ex.ParamName)
            Console.WriteLine(ex.Message.Length > 0)
        End Try
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["value", "True"]);
}

#[test]
fn exception_argument_null_exception_tracks_argument_name() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Try
            Throw New ArgumentNullException("payload")
        Catch ex As ArgumentNullException
            Console.WriteLine(ex.ParamName)
            Console.WriteLine(ex.GetType().Name)
        End Try
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["payload", "ArgumentNullException"]);
}

#[test]
fn exception_argument_out_of_range_exposes_actual_value() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Try
            Throw New ArgumentOutOfRangeException("index", 13, "value must be non-negative")
        Catch ex As ArgumentOutOfRangeException
            Console.WriteLine(ex.ParamName)
            Console.WriteLine(ex.ActualValue = 13)
            Console.WriteLine(ex.GetType().Name)
        End Try
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["index", "True", "ArgumentOutOfRangeException"]);
}

#[test]
fn exception_format_exception_from_integer_parse() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Try
            Integer.Parse("not-a-number")
        Catch ex As FormatException
            Console.WriteLine(ex.GetType().Name)
        End Try
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["FormatException"]);
}

#[test]
fn exception_invalid_cast_exception_from_bad_conversion() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Try
            Dim value As Object = "text"
            Dim typed As Integer = CInt(value)
            Console.WriteLine(typed)
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
fn exception_key_not_found_is_thrown_for_missing_dictionary_key() {
    let out = run_vb(
        r#"
Imports System
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)()
        map("present") = 11

        Try
            Console.WriteLine(map("missing"))
        Catch ex As KeyNotFoundException
            Console.WriteLine(ex.GetType().Name)
        End Try
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["KeyNotFoundException"]);
}

#[test]
fn exception_not_implemented_is_detectable_by_type() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Try
            Throw New NotImplementedException()
        Catch ex As NotImplementedException
            Console.WriteLine(ex.GetType().Name)
        End Try
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["NotImplementedException"]);
}

#[test]
fn exception_object_disposed_exposes_object_name() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Try
            Throw New ObjectDisposedException("MyObject")
        Catch ex As ObjectDisposedException
            Console.WriteLine(ex.ObjectName)
        End Try
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["MyObject"]);
}

#[test]
fn exception_uri_format_exception_from_bad_uri() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Try
            Dim url As New Uri("://definitely-not-valid")
            Console.WriteLine(url.AbsoluteUri)
        Catch ex As UriFormatException
            Console.WriteLine(ex.GetType().Name)
        End Try
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["UriFormatException"]);
}

#[test]
fn exception_overflow_exception_from_narrowing_conversion() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Try
            Dim value As Byte = CByte(300)
            Console.WriteLine(value)
        Catch ex As OverflowException
            Console.WriteLine(ex.GetType().Name)
        End Try
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["OverflowException"]);
}

#[test]
fn exception_throws_can_carry_inner_exception_messages() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Try
            Try
                Throw New Exception("root")
            Catch inner As Exception
                Throw New Exception("wrapper", inner)
            End Try
        Catch ex As Exception
            Console.WriteLine(ex.InnerException IsNot Nothing)
            Console.WriteLine(ex.InnerException.Message)
        End Try
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "root"]);
}
