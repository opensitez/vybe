' vybe-test: vb/vb_named_tuples_operations/test_vb_named_tuple_inference_from_variable_names
' origin: languages/vb/tests/vb/test_vb_named_tuples_operations.rs

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
    Sub Main()
        Dim count As Integer = 5
        Dim label As String = "Total"
        Dim t = (count, label)
        __Check(CStr(t.count & ":" & t.label), "5:Total")
    End Sub
End Module
