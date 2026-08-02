' vybe-test: vb/vb_object_late_bound_method_call/test_vb_late_bound_method_call_multiple_args
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
    Class TextFormatter
        Public Function ConcatStrings(a As String, b As String, c As String) As String
            Return a & "-" & b & "-" & c
        End Function
    End Class

    Sub Main()
        Dim obj As Object = New TextFormatter()
        Dim res As String = CStr(obj.ConcatStrings("A", "B", "C"))
        __Check(CStr(res), "A-B-C")
    End Sub
End Module
