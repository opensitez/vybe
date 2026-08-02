' vybe-test: vb/vb_iif_ternary_evaluation/test_vb_if_operator_null_reference_method_guard
' origin: languages/vb/tests/vb/test_vb_iif_ternary_evaluation.rs

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

Class Customer
    Public Property Name As String = "Alice"
End Class

Module Program
    Sub Main()
        Dim c As Customer = Nothing
        Dim name = If(c IsNot Nothing, c.Name, "Unknown")
        __Check(CStr(name), "Unknown")
    End Sub
End Module
