' vybe-test: vb/vb_tuple_deconstruct_method_overloads/test_vb_tuple_byref_argument_passing
' origin: languages/vb/tests/vb/test_vb_tuple_deconstruct_method_overloads.rs

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
    Private Sub ModifyTuple(ByRef t As (Integer, String))
        t.Item1 = 99
        t.Item2 = "Updated"
    End Sub

    Sub Main()
        Dim t = (10, "Original")
        ModifyTuple(t)
        __Check(CStr(t.Item1 & ":" & t.Item2), "99:Updated")
    End Sub
End Module
