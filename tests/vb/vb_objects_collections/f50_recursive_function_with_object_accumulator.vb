' vybe-test: vb/vb_objects_collections/f50_recursive_function_with_object_accumulator
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

Class Acc
    Public Total As Integer
End Class
Sub AddUp(a As Acc, n As Integer)
    If n <= 0 Then Return
    a.Total = a.Total + n
    AddUp(a, n - 1)
End Sub
Dim acc As New Acc()
acc.Total = 0
AddUp(acc, 5)
__Check(CStr(acc.Total), "15")
