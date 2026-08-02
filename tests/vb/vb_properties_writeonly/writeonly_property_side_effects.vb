' vybe-test: vb/vb_properties_writeonly/writeonly_property_side_effects
' origin: languages/vb/tests/vb/test_vb_properties_writeonly.rs

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

Class Counter
    Public Total As Integer = 0
    
    Public WriteOnly Property AddAmount As Integer
        Set(value As Integer)
            Total = Total + value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim c As New Counter()
        c.AddAmount = 5
        c.AddAmount = 10
        __Check(CStr(c.Total), "15")
    End Sub
End Module
