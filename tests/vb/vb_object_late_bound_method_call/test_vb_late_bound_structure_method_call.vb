' vybe-test: vb/vb_object_late_bound_method_call/test_vb_late_bound_structure_method_call
' origin: languages/vb/tests/vb/test_vb_object_late_bound_method_call.rs

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
    Structure Vector2D
        Public X As Integer
        Public Y As Integer
        Public Function GetMagnitude() As Double
            Return Math.Sqrt(X * X + Y * Y)
        End Function
    End Structure

    Sub Main()
        Dim obj As Object = New Vector2D With {.X = 3, .Y = 4}
        __Check(CStr(CDbl(obj.GetMagnitude())), "5")
    End Sub
End Module
