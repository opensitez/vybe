' vybe-test: vb/vb_optional_paramarrays/paramarray_and_optional_together
' origin: languages/vb/tests/vb/test_vb_optional_paramarrays.rs

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
' Function F(Optional a As Integer = 1, ParamArray args() As Integer) ' Cannot combine Optional and ParamArray
Sub Main()
__Check(CStr("Parsed"), "Parsed")
End Sub
End Module
