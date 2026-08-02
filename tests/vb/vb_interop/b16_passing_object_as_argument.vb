' vybe-test: vb/vb_interop/b16_passing_object_as_argument
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
    Public Function GetName() As String
        Return name
    End Function
End Class
Function Describe(item As Object) As String
    Return "Item: " & item.GetName()
End Function
Dim it As New Item("Widget")
__Check(CStr(Describe(it)), "Item: Widget")
