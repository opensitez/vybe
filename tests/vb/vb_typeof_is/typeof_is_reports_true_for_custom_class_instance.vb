' vybe-test: vb/vb_typeof_is/typeof_is_reports_true_for_custom_class_instance
' origin: languages/vb/tests/vb/test_vb_typeof_is.rs

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

Module M
    Class Greeter
    End Class

    Sub Main()
        Dim obj As Object = New Greeter()
        __Check(CStr(TypeOf obj Is Greeter), "True")
    End Sub
End Module
