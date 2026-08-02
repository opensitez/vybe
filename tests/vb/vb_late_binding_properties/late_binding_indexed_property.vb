' vybe-test: vb/vb_late_binding_properties/late_binding_indexed_property
' origin: languages/vb/tests/vb/test_vb_late_binding_properties.rs

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

Class ItemCollection
    Private items As String() = {"A", "B", "C"}
    
    Default Public Property Item(index As Integer) As String
        Get
            Return items(index)
        End Get
        Set(value As String)
            items(index) = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim col As Object = New ItemCollection()
        
        ' Late bound indexed property via Default property (implicit)
        col(1) = "Z"
        
        ' Late bound indexed property (explicit)
        __Check(CStr(col.Item(0)), "A")
        __Check(CStr(col(1)), "Z")
    End Sub
End Module
