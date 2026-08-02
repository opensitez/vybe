' vybe-test: vb/vb_byref_mutation/byref_can_mutate_structure_fields_in_place
' origin: languages/vb/tests/vb/test_vb_byref_mutation.rs

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
    Structure Point
        Public X As Integer
        Public Y As Integer
    End Structure

    Sub Move(ByRef point As Point)
        point.X += 1
        point.Y += 2
    End Sub

    Sub Main()
        Dim point As Point
        point.X = 2
        point.Y = 3
        Move(point)
        __Check(CStr(point.X.ToString() & "," & point.Y.ToString()), "3,5")
    End Sub
End Module
