' vybe-test: vb/vb_spec_object_model/object_model_spec_default_property_can_index_collection_like_type
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

Class Buffer
    Private values() As String = {"a", "b", "c"}
    Default Public Property Item(index As Integer) As String
        Get
            Return values(index)
        End Get
        Set(value As String)
            values(index) = value
        End Set
    End Property
End Class
Module M
    Sub Main()
        Dim buffer As New Buffer()
        buffer(1) = "x"
        __Check(CStr(buffer(1)), "x")
    End Sub
End Module
