use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Custom Exception Class Inheritance
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_custom_exception_class_properties() {
    let src = r#"
Imports System

Class ValidationException
    Inherits Exception
    Public ReadOnly FieldName As String

    Public Sub New(fieldName As String, message As String)
        MyBase.New(message)
        Me.FieldName = fieldName
    End Sub
End Class

Module Program
    Sub Main()
        Try
            Throw New ValidationException("Email", "Invalid email address format")
        Catch ex As ValidationException
            Console.WriteLine(ex.FieldName & ":" & ex.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Email:Invalid email address format"]);
}

#[test]
fn test_vb_custom_exception_constructors_standard() {
    let src = r#"
Imports System

Class CustomNotFoundException
    Inherits Exception

    Public Sub New()
        MyBase.New("Resource not found")
    End Sub

    Public Sub New(message As String)
        MyBase.New(message)
    End Sub

    Public Sub New(message As String, inner As Exception)
        MyBase.New(message, inner)
    End Sub
End Class

Module Program
    Sub Main()
        Try
            Try
                Throw New InvalidOperationException("Inner error")
            Catch innerEx As Exception
                Throw New CustomNotFoundException("Outer error", innerEx)
            End Try
        Catch outerEx As CustomNotFoundException
            Console.WriteLine(outerEx.Message)
            Console.WriteLine(outerEx.InnerException.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Outer error", "Inner error"]);
}
