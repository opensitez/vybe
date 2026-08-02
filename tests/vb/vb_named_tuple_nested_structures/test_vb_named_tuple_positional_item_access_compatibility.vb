' vybe-test: vb/vb_named_tuple_nested_structures/test_vb_named_tuple_positional_item_access_compatibility
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
        Dim pt As (X As Integer, Y As Integer) = (10, 20)
        ' Positional .Item1 and .Item2 are aliases for .X and .Y
        __Check(CStr(pt.Item1 & "," & pt.Item2), "10,20")
    End Sub
End Module
