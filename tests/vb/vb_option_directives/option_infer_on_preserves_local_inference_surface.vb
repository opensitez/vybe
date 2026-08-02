' vybe-test: vb/vb_option_directives/option_infer_on_preserves_local_inference_surface
' origin: languages/vb/tests/vb/test_vb_option_directives.rs

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

Option Infer On
Module M
    Sub Main()
        Dim total = 6
        __Check(CStr(total + 1), "7")
    End Sub
End Module
