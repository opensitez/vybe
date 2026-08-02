' vybe-test: vb/vb_interop/a03_method_accesses_field_without_me_prefix
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

Public Class Person
    Dim name As String
    Dim age As Integer
    Public Sub New(n As String, a As Integer)
        name = n
        age = a
    End Sub
    Public Function Describe() As String
        Return name & " is " & CStr(age)
    End Function
End Class
Dim p As New Person("Alice", 30)
__Check(CStr(p.Describe()), "Alice is 30")
