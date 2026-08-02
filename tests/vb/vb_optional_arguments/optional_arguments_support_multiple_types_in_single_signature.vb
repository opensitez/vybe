' vybe-test: vb/vb_optional_arguments/optional_arguments_support_multiple_types_in_single_signature
' origin: languages/vb/tests/vb/test_vb_optional_arguments.rs

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
    Function Describe(name As String, Optional level As Integer = 1, Optional suffix As String = "ok") As String
        Return name & ":" & level & ":" & suffix
    End Function

    Sub Main()
        __Check(CStr(Describe("Hope")), "Hope:1:ok")
        __Check(CStr(Describe("Ivy", 3, "done")), "Ivy:3:done")
    End Sub
End Module
