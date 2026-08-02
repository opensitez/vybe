' vybe-test: vb/vb_interop/f49_object_in_array_access_via_index
' origin: languages/vb/tests/vb/vb_interop_test.rs

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

Public Class Item
    Dim name As String
    Public Sub New(n As String)
        name = n
    End Sub
End Class
Dim items(2) As Item
items(0) = New Item("first")
items(1) = New Item("second")
items(2) = New Item("third")
__Check(CStr(items(0).name), "first")
__Check(CStr(items(2).name), "third")
