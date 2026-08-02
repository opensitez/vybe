' vybe-test: vb/vb_exit_property_set/exit_property_set
' origin: languages/vb/tests/vb/test_vb_exit_property_set.rs

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

Class Cache
    Private _val As Integer
    
    Public Property Value As Integer
        Get
            Return _val
        End Get
        Set(val As Integer)
            If val < 0 Then Exit Property
            _val = val
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim c As New Cache()
        c.Value = -10
        __Check(CStr(c.Value), "0") ' Should be 0
        c.Value = 20
        __Check(CStr(c.Value), "20") ' Should be 20
    End Sub
End Module
