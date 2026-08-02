' vybe-test: vb/vb_named_arguments/named_arguments_can_use_function_results_as_values
' origin: languages/vb/tests/vb/test_vb_named_arguments.rs

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
    Function Prefix() As String
        Return "pre"
    End Function

    Function Suffix() As String
        Return "post"
    End Function

    Function Describe(name As String, prefix As String, suffix As String) As String
        Return prefix & ":" & name & ":" & suffix
    End Function

    Sub Main()
        __Check(CStr(Describe(name:="Hope", suffix:=Suffix(), prefix:=Prefix())), "pre:Hope:post")
    End Sub
End Module
