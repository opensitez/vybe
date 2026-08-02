' vybe-test: vb/vb_mod_operator_floating_point/test_vb_mod_operator_even_odd_checker
' origin: languages/vb/tests/vb/test_vb_mod_operator_floating_point.rs

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
    Private Function IsEven(n As Integer) As Boolean
        Return (n Mod 2) = 0
    End Function

    Sub Main()
        __Check(CStr(IsEven(10) & "|" & IsEven(11)), "True|False")
    End Sub
End Module
