use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Generics (Constraints)
// ═══════════════════════════════════════════════════════════

#[test]
fn generic_constraint_new() {
    let out = run_vb(
        r#"
Class Factory(Of T As New)
    Public Function CreateInstance() As T
        Return New T()
    End Function
End Class

Class Widget
    Public Name As String = "DefaultWidget"
End Class

Module M
    Sub Main()
        Dim f As New Factory(Of Widget)()
        Dim w As Widget = f.CreateInstance()
        Console.WriteLine(w.Name)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["DefaultWidget"]);
}

#[test]
fn generic_constraint_class_and_interface() {
    let out = run_vb(
        r#"
Interface IIdentifiable
    Function GetID() As Integer
End Interface

' T must be a reference type and implement IIdentifiable
Class Repository(Of T As {Class, IIdentifiable})
    Public Sub Process(item As T)
        Console.WriteLine("Processing ID: " & item.GetID().ToString())
    End Sub
End Class

Class User
    Implements IIdentifiable
    Public Function GetID() As Integer Implements IIdentifiable.GetID
        Return 999
    End Function
End Class

Module M
    Sub Main()
        Dim r As New Repository(Of User)()
        Dim u As New User()
        r.Process(u)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Processing ID: 999"]);
}

#[test]
fn generic_constraint_structure() {
    let out = run_vb(
        r#"
Module M
    ' T must be a value type
    Function GetDefault(Of T As Structure)() As T
        Dim temp As T
        Return temp
    End Function

    Sub Main()
        Dim d As Integer = GetDefault(Of Integer)()
        Console.WriteLine(d)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["0"]);
}
