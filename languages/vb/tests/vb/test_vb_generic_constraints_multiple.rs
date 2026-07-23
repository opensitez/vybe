use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Multiple Generic Type Constraints (Class, Structure, New, Interfaces)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_generic_constraint_class_and_new() {
    let src = r#"
Imports System

Class Factory(Of T As {Class, New})
    Public Function CreateInstance() As T
        Return New T()
    End Function
End Class

Class Item
    Public Tag As String = "Created"
End Class

Module Program
    Sub Main()
        Dim f As New Factory(Of Item)()
        Dim i As Item = f.CreateInstance()
        Console.WriteLine(i.Tag)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Created"]);
}

#[test]
fn test_vb_generic_constraint_multiple_interfaces() {
    let src = r#"
Imports System
Imports System.Collections

Interface IIdentifiable
    ReadOnly Property Id As Integer
End Interface

Class EntityContainer(Of T As {IIdentifiable, IComparable})
    Public Sub Process(item As T)
        Console.WriteLine("Id: " & item.Id)
    End Sub
End Class

Class Product
    Implements IIdentifiable, IComparable
    Public ReadOnly Property Id As Integer Implements IIdentifiable.Id
        Get
            Return 100
        End Get
    End Property

    Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
        Return 0
    End Function
End Class

Module Program
    Sub Main()
        Dim container As New EntityContainer(Of Product)()
        container.Process(New Product())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Id: 100"]);
}

#[test]
fn test_vb_generic_constraint_structure_value_type() {
    let src = r#"
Imports System

Class ValueHolder(Of T As Structure)
    Public Value As T
    Public Sub New(v As T)
        Me.Value = v
    End Sub
End Class

Module Program
    Sub Main()
        Dim h As New ValueHolder(Of Integer)(42)
        Console.WriteLine(h.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42"]);
}
