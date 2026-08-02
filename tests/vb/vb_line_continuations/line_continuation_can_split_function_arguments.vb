' vybe-test: vb/vb_line_continuations/line_continuation_can_split_function_arguments
' origin: languages/vb/tests/vb/test_vb_line_continuations.rs

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
    Function Add(a As Integer, b As Integer, c As Integer) As Integer
        Return a + b + c
    End Function

    Sub Main()
        __Check(CStr(Add(1, _
            3, _
            5)), "9")
    End Sub
End Module
