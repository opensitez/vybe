' vybe-test: vb/vb_objects_collections/e41_array_of_objects
' origin: languages/vb/tests/vb/vb_objects_collections_test.rs

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

Class Pt
    Public X As Integer
    Public Y As Integer
End Class
Dim pts(3) As Pt
Dim p As New Pt()
p.X = 5
p.Y = 10
pts(0) = p
__Check(CStr(pts(0).X), "5")
__Check(CStr(pts(0).Y), "10")
