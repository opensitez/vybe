use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: NullReferenceException & Null Safety Guards (Nothing)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_null_instance_method_call_throws_null_reference() {
    let src = r#"
Imports System

Class Document
    Public Sub Print() : Console.WriteLine("Print") : End Sub
End Class

Module Program
    Sub Main()
        Dim doc As Document = Nothing
        Try
            doc.Print()
        Catch ex As NullReferenceException
            Console.WriteLine("NullReferenceException Caught on Method Call")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["NullReferenceException Caught on Method Call"]
    );
}

#[test]
fn test_vb_null_instance_property_getter_throws_null_reference() {
    let src = r#"
Imports System

Class User
    Public Property Name As String
End Class

Module Program
    Sub Main()
        Dim u As User = Nothing
        Try
            Dim n = u.Name
            Console.WriteLine(n)
        Catch ex As NullReferenceException
            Console.WriteLine("NullReferenceException Caught on Property Get")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["NullReferenceException Caught on Property Get"]
    );
}

#[test]
fn test_vb_null_instance_property_setter_throws_null_reference() {
    let src = r#"
Imports System

Class User
    Public Property Name As String
End Class

Module Program
    Sub Main()
        Dim u As User = Nothing
        Try
            u.Name = "Alice"
        Catch ex As NullReferenceException
            Console.WriteLine("NullReferenceException Caught on Property Set")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["NullReferenceException Caught on Property Set"]
    );
}

#[test]
fn test_vb_null_conditional_operator_method_call() {
    let src = r#"
Class Document
    Public Function GetTitle() As String
        Return "ValidTitle"
    End Function
End Class

Module Program
    Sub Main()
        Dim doc As Document = Nothing
        Dim title As String = doc?.GetTitle()
        Console.WriteLine(title Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_null_conditional_operator_property_access() {
    let src = r#"
Class Address
    Public Property City As String = "Seattle"
End Class

Class Person
    Public Property HomeAddress As Address
End Class

Module Program
    Sub Main()
        Dim p As New Person()
        Console.WriteLine(p.HomeAddress?.City Is Nothing)
        p.HomeAddress = New Address()
        Console.WriteLine(p.HomeAddress?.City)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "Seattle"]);
}

#[test]
fn test_vb_null_coalescing_if_function() {
    let src = r#"
Module Program
    Sub Main()
        Dim name As String = Nothing
        Dim displayName As String = If(name, "Guest")
        Console.WriteLine(displayName)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Guest"]);
}

#[test]
fn test_vb_null_coalescing_if_function_non_null() {
    let src = r#"
Module Program
    Sub Main()
        Dim name As String = "Alice"
        Dim displayName As String = If(name, "Guest")
        Console.WriteLine(displayName)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice"]);
}

#[test]
fn test_vb_null_conditional_indexer_access() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As String() = Nothing
        Dim first = arr?(0)
        Console.WriteLine(first Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_is_nothing_comparison() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = Nothing
        Console.WriteLine(s Is Nothing)
        Console.WriteLine(IsNothing(s))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "True"]);
}

#[test]
fn test_vb_isnot_nothing_comparison() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "Hello"
        Console.WriteLine(s IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_null_reference_field_access() {
    let src = r#"
Imports System

Class Container
    Public Data As String = "Value"
End Class

Module Program
    Sub Main()
        Dim c As Container = Nothing
        Try
            Dim d = c.Data
        Catch ex As NullReferenceException
            Console.WriteLine("Field Access NullReferenceException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Field Access NullReferenceException Caught"]
    );
}

#[test]
fn test_vb_null_reference_event_raise_guard() {
    let src = r#"
Imports System

Class Publisher
    Public Event Notify As Action
    Public Sub Fire()
        ' RaiseEvent in VB automatically guards against null internal delegate!
        RaiseEvent Notify()
        Console.WriteLine("Fire executed safely without subscribers")
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New Publisher()
        p.Fire()
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Fire executed safely without subscribers"]
    );
}

#[test]
fn test_vb_nullable_value_type_has_value_check() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim num As Nullable(Of Integer) = Nothing
        If Not num.HasValue Then
            Console.WriteLine("Value is Nothing")
        End If
        Console.WriteLine(num.GetValueOrDefault(100))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Value is Nothing", "100"]);
}

#[test]
fn test_vb_nullable_value_type_value_property_throws_invalid_operation() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim num As Nullable(Of Integer) = Nothing
        Try
            Dim val = num.Value
        Catch ex As InvalidOperationException
            Console.WriteLine("Nullable.Value InvalidOperationException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Nullable.Value InvalidOperationException Caught"]
    );
}

#[test]
fn test_vb_chained_null_conditional_calls() {
    let src = r#"
Class Company
    Public Property Owner As Person
End Class

Class Person
    Public Property Address As Address
End Class

Class Address
    Public Property ZipCode As String = "90210"
End Class

Module Program
    Sub Main()
        Dim comp As Company = Nothing
        Console.WriteLine(comp?.Owner?.Address?.ZipCode Is Nothing)
        comp = New Company() With {.Owner = New Person() With {.Address = New Address()}}
        Console.WriteLine(comp?.Owner?.Address?.ZipCode)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "90210"]);
}

#[test]
fn test_vb_argument_null_exception_guard() {
    let src = r#"
Imports System

Module Program
    Private Sub ProcessData(data As String)
        If data Is Nothing Then
            Throw New ArgumentNullException(NameOf(data), "Data cannot be null")
        End If
    End Sub

    Sub Main()
        Try
            ProcessData(Nothing)
        Catch ex As ArgumentNullException
            Console.WriteLine("Param: " & ex.ParamName)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Param: data"]);
}

#[test]
fn test_vb_null_array_length_access_throws_null_reference() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim arr As Integer() = Nothing
        Try
            Dim len = arr.Length
        Catch ex As NullReferenceException
            Console.WriteLine("Null Array Length NullReferenceException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Null Array Length NullReferenceException Caught"]
    );
}

#[test]
fn test_vb_null_conditional_delegate_invocation() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim action As Action = Nothing
        action?.Invoke()
        Console.WriteLine("Null Action?.Invoke() executed safely")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Null Action?.Invoke() executed safely"]);
}

#[test]
fn test_vb_string_isnullorwhitespace_guard() {
    let src = r#"
Module Program
    Sub Main()
        Dim s1 As String = Nothing
        Dim s2 As String = "   "
        Dim s3 As String = "VB"
        Console.WriteLine(String.IsNullOrWhiteSpace(s1) & "|" & String.IsNullOrWhiteSpace(s2) & "|" & String.IsNullOrWhiteSpace(s3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|False"]);
}

#[test]
fn test_vb_null_unboxing_cast_throws_null_reference() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim obj As Object = Nothing
        Try
            Dim i As Integer = CInt(obj)
            Console.WriteLine(i)
        Catch ex As NullReferenceException
            Console.WriteLine("Unboxing Null NullReferenceException Caught")
        Catch ex As Exception
            Console.WriteLine("Caught: " & ex.GetType().Name)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Caught: NullReferenceException"]);
}
