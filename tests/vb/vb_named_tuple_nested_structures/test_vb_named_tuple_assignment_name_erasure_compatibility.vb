' vybe-test: vb/vb_named_tuple_nested_structures/test_vb_named_tuple_assignment_name_erasure_compatibility
' origin: languages/vb/tests/vb/test_vb_named_tuple_nested_structures.rs

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
        Dim t1 As (Width As Integer, Height As Integer) = (100, 200)
        Dim t2 As (W As Integer, H As Integer) = t1 ' Names are erased at runtime; underlying ValueTuple is compatible!
        __Check(CStr(t2.W & "x" & t2.H), "100x200")
    End Sub
End Module
