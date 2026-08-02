' vybe-test: vb/vb_parser_traps/property_get_set_different_access
' origin: languages/vb/tests/vb/test_vb_parser_traps.rs

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

Class Data
    Private _val As Integer
    
    ' Only one accessor can have an access modifier different from the property
    Public Property Val As Integer
        Get
            Return _val
        End Get
        Protected Set(value As Integer)
            _val = value
        End Set
    End Property
    
    Public Sub Update(v As Integer)
        Val = v
    End Sub
End Class

Module M
    Sub Main()
        Dim d As New Data()
        d.Update(100)
        __Check(CStr(d.Val), "100")
    End Sub
End Module
