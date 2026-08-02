' vybe-test: vb/vb_objects_collections/f45_sub_modifies_object_field
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

Class Score
    Public Points As Integer
End Class
Sub AddPoints(s As Score, p As Integer)
    s.Points = s.Points + p
End Sub
Dim sc As New Score()
sc.Points = 100
AddPoints(sc, 50)
__Check(CStr(sc.Points), "150")
