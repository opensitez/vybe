' vybe-test: vb/vb_default_property_multi_index/default_property_multi_index
' origin: languages/vb/tests/vb/test_vb_default_property_multi_index.rs

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

Class Matrix
    Private data(10, 10) As Integer
    
    Default Public Property Item(x As Integer, y As Integer) As Integer
        Get
            Return data(x, y)
        End Get
        Set(value As Integer)
            data(x, y) = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim m As New Matrix()
        m(5, 5) = 42
        __Check(CStr(m(5, 5)), "42")
    End Sub
End Module
