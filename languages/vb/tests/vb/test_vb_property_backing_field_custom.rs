use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Custom Property Getters/Setters & Backing Fields
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_property_validation_in_setter() {
    let src = r#"
Class AgeTracker
    Private _age As Integer

    Public Property Age As Integer
        Get
            Return _age
        End Get
        Set(value As Integer)
            If value < 0 Then
                _age = 0
            Else
                _age = value
            End If
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim t As New AgeTracker()
        t.Age = -5
        Console.WriteLine(t.Age)
        t.Age = 25
        Console.WriteLine(t.Age)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0", "25"]);
}

#[test]
fn test_vb_property_lazy_backing_field() {
    let src = r#"
Class LazyData
    Private _data As String = Nothing

    Public ReadOnly Property Data As String
        Get
            If _data Is Nothing Then
                _data = "ComputedValue"
            End If
            Return _data
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim d As New LazyData()
        Console.WriteLine(d.Data)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ComputedValue"]);
}

#[test]
fn test_vb_property_change_notification_backing_field() {
    let src = r#"
Class NotifyingItem
    Private _val As Integer
    Public Event ValueChanged(oldV As Integer, newV As Integer)

    Public Property Value As Integer
        Get
            Return _val
        End Get
        Set(val As Integer)
            If _val <> val Then
                Dim old As Integer = _val
                _val = val
                RaiseEvent ValueChanged(old, val)
            End If
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim item As New NotifyingItem()
        AddHandler item.ValueChanged, Sub(oldV, newV)
            Console.WriteLine("Changed: " & oldV & "->" & newV)
        End Sub
        item.Value = 10
        item.Value = 10 ' No event
        item.Value = 20
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Changed: 0->10", "Changed: 10->20"]);
}
