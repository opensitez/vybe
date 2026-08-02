' vybe-test: vb/vb_property_backing_field_custom/test_vb_property_change_notification_backing_field
' origin: languages/vb/tests/vb/test_vb_property_backing_field_custom.rs

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
