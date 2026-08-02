' vybe-test: vb/vb_default_property_generic/default_property_generic
' origin: languages/vb/tests/vb/test_vb_default_property_generic.rs

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

Class Cache(Of T)
    Private _dict As New System.Collections.Generic.Dictionary(Of String, T)()
    
    Default Public Property Item(key As String) As T
        Get
            If _dict.ContainsKey(key) Then Return _dict(key)
            Return Nothing
        End Get
        Set(value As T)
            _dict(key) = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim c As New Cache(Of Integer)()
        c("A") = 100
        __Check(CStr(c("A")), "100")
    End Sub
End Module
