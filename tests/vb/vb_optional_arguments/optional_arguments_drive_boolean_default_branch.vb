' vybe-test: vb/vb_optional_arguments/optional_arguments_drive_boolean_default_branch
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
    Function Render(label As String, Optional loud As Boolean = False) As String
        If loud Then
            Return label & "!"
        End If
        Return label & "."
    End Function

    Sub Main()
        __Check(CStr(Render("calm")), "calm.")
    End Sub
End Module
