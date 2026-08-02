' vybe-test: vb/vb_spec_delegates_lambdas/delegate_spec_delegate_can_invoke_instance_method_after_state_change
' origin: languages/vb/tests/vb/test_vb_spec_delegates_lambdas.rs

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
    Public Value As Integer
    Public Function Read() As Integer
        Return Value
    End Function
End Class
Module M
    Sub Main()
        Dim c As New Counter()
        Dim fn As Func(Of Integer) = AddressOf c.Read
        c.Value = 11
        __Check(CStr(fn()), "11")
    End Sub
End Module
