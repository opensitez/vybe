' vybe-test: vb/vb_named_tuple_nested_structures/test_vb_nested_tuple_structure_access
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
        Dim line As (StartPt As (X As Integer, Y As Integer), EndPt As (X As Integer, Y As Integer))
        line.StartPt = (0, 0)
        line.EndPt = (10, 20)
        __Check(CStr(line.StartPt.X & "," & line.StartPt.Y & " -> " & line.EndPt.X & "," & line.EndPt.Y), "0,0 -> 10,20")
    End Sub
End Module
