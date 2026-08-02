' vybe-test: vb/vb_structs_basic/struct_with_block
' origin: languages/vb/tests/vb/test_vb_structs_basic.rs

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

Structure Rect
    Public Width As Integer
    Public Height As Integer
End Structure

Module M
    Sub Main()
        Dim r As Rect
        With r
            .Width = 100
            .Height = 200
        End With
        __Check(CStr(r.Width), "100")
    End Sub
End Module
