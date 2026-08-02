' vybe-test: vb/vb_object_late_bound_method_call/test_vb_late_bound_method_call_single_arg
' origin: languages/vb/tests/vb/test_vb_object_late_bound_method_call.rs

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

Module Program
    Class Calculator
        Public Function AddTen(x As Integer) As Integer
            Return x + 10
        End Function
    End Class

    Sub Main()
        Dim obj As Object = New Calculator()
        Dim res As Integer = CInt(obj.AddTen(5))
        __Check(CStr(res), "15")
    End Sub
End Module
