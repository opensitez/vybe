' vybe-test: vb/vb_named_tuples_operations/test_vb_named_tuple_literal_access
' origin: languages/vb/tests/vb/test_vb_named_tuples_operations.rs

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
        Dim item As (Id As Integer, Name As String) = (1, "Widget")
        __Check(CStr(item.Id & ":" & item.Name), "1:Widget")
        __Check(CStr(item.Item1 & ":" & item.Item2), "1:Widget")
    End Sub
End Module
