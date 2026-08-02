' vybe-test: vb/vb_optional_arguments/optional_arguments_can_chain_through_helper_functions
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
    Function Decorate(name As String, Optional prefix As String = "base", Optional suffix As String = ".") As String
        Return prefix & ":" & name & ":" & suffix
    End Function

    Function Outer(name As String, Optional prefix As String = "outer") As String
        Return Decorate(name, prefix)
    End Function

    Sub Main()
        __Check(CStr(Outer("Faye")), "outer:Faye:.")
        __Check(CStr(Outer("Gus", "inner")), "inner:Gus:.")
    End Sub
End Module
