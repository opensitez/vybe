' vybe-test: vb/vb_structs_nested/structs_nested_initializers
' origin: languages/vb/tests/vb/test_vb_structs_nested.rs

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

Structure Point
    Public X As Integer
    Public Y As Integer
End Structure

Structure Rectangle
    Public TopLeft As Point
    Public BottomRight As Point
End Structure

Module M
    Sub Main()
        Dim r As New Rectangle() With {
            .TopLeft = New Point() With {.X = 0, .Y = 10},
            .BottomRight = New Point() With {.X = 10, .Y = 0}
        }
        __Check(CStr(r.TopLeft.Y), "10")
        __Check(CStr(r.BottomRight.X), "10")
    End Sub
End Module
