' vybe-test: vb/vb_event_delegate_signature_matching/test_vb_delegate_custom_signature_function
' origin: languages/vb/tests/vb/test_vb_event_delegate_signature_matching.rs

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

Delegate Function Transform(input As String) As String

Module Program
    Private Function Upper(s As String) As String
        Return s.ToUpper()
    End Function

    Sub Main()
        Dim t As Transform = AddressOf Upper
        __Check(CStr(t("hello")), "HELLO")
    End Sub
End Module
