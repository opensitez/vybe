' vybe-test: vb/vb_tuple_deconstruction_pattern/test_vb_tuple_deconstruct_into_existing_vars
' origin: languages/vb/tests/vb/test_vb_tuple_deconstruction_pattern.rs

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
        Dim t = (10, "Ten")
        Dim x As Integer
        Dim y As String
        (x, y) = t
        __Check(CStr(x & ":" & y), "10:Ten")
    End Sub
End Module
