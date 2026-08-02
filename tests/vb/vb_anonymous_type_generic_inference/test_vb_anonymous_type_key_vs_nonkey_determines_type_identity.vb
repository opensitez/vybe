' vybe-test: vb/vb_anonymous_type_generic_inference/test_vb_anonymous_type_key_vs_nonkey_determines_type_identity
' origin: languages/vb/tests/vb/test_vb_anonymous_type_generic_inference.rs

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
        Dim o1 = New With {Key .A = 1, .B = "X"}
        Dim o2 = New With {.A = 1, .B = "X"}
        __Check(CStr(o1.GetType() Is o2.GetType()), "False")
    End Sub
End Module
