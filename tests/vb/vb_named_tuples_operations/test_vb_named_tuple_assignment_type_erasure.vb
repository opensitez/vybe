' vybe-test: vb/vb_named_tuples_operations/test_vb_named_tuple_assignment_type_erasure
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
        Dim t1 As (A As Integer, B As String) = (10, "Ten")
        Dim t2 As (X As Integer, Y As String) = t1 ' Type name erase assignment
        __Check(CStr(t2.X & ":" & t2.Y), "10:Ten")
    End Sub
End Module
