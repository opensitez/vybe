' vybe-test: vb/vb_iif_ternary_evaluation/test_vb_if_operator_structure_result
' origin: languages/vb/tests/vb/test_vb_iif_ternary_evaluation.rs

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

Structure Point2D
    Public X, Y As Integer
End Structure

Module Program
    Sub Main()
        Dim p1 As New Point2D With {.X = 1, .Y = 1}
        Dim p2 As New Point2D With {.X = 2, .Y = 2}
        Dim selected = If(True, p1, p2)
        __Check(CStr(selected.X & "," & selected.Y), "1,1")
    End Sub
End Module
