' vybe-test: vb/vb_spec_delegates_lambdas/delegate_spec_addressof_instance_method_matches_func_delegate
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
    Public Function DoubleValue(x As Integer) As Integer
        Return x * 2
    End Function
End Class
Module M
    Sub Main()
        Dim c As New Counter()
        Dim fn As Func(Of Integer, Integer) = AddressOf c.DoubleValue
        __Check(CStr(fn(8)), "16")
    End Sub
End Module
