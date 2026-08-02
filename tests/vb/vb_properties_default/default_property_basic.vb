' vybe-test: vb/vb_properties_default/default_property_basic
' origin: languages/vb/tests/vb/test_vb_properties_default.rs

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

Class StringCollection
    Private _items(10) As String
    
    Default Public Property Item(index As Integer) As String
        Get
            Return _items(index)
        End Get
        Set(value As String)
            _items(index) = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim col As New StringCollection()
        ' Accessing via default property syntax
        col(0) = "First"
        col(1) = "Second"
        
        __Check(CStr(col(0)), "First")
        __Check(CStr(col(1)), "Second")
    End Sub
End Module
