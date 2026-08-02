' vybe-test: vb/vb_isnot_operator_null_checks/test_vb_isnot_operator_generic_class_type_check
' origin: languages/vb/tests/vb/test_vb_isnot_operator_null_checks.rs

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

Class Wrapper(Of T)
    Public Item As T
End Class

Module Program
    Sub Main()
        Dim w As New Wrapper(Of String) With {.Item = "Data"}
        __Check(CStr(w IsNot Nothing AndAlso w.Item IsNot Nothing), "True")
    End Sub
End Module
