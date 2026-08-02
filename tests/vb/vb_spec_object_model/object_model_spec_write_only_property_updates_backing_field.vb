' vybe-test: vb/vb_spec_object_model/object_model_spec_write_only_property_updates_backing_field
' origin: languages/vb/tests/vb/test_vb_spec_object_model.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Class Bag
    Private _count As Integer
    Public WriteOnly Property Count As Integer
        Set(value As Integer)
            _count = value
        End Set
    End Property
    Public Function Snapshot() As Integer
        Return _count
    End Function
End Class
Module M
    Sub Main()
        Dim b As New Bag()
        b.Count = 5
        __Check(CStr(b.Snapshot()), "5")
    End Sub
End Module
