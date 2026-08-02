' vybe-test: vb/vb_is_operator_runtime_type_check/test_vb_is_operator_string_non_interned_reference_check
' origin: languages/vb/tests/vb/test_vb_is_operator_runtime_type_check.rs

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
        Dim s1 As String = New String({"A"c, "B"c})
        Dim s2 As String = New String({"A"c, "B"c})
        __Check(CStr((s1 = s2) & "|" & (s1 Is s2)), "True|False")
    End Sub
End Module
