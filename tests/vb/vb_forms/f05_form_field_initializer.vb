' vybe-test: vb/vb_forms/f05_form_field_initializer
' origin: languages/vb/tests/vb/vb_forms_test.rs

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

Public Class Form1
    Dim greeting As String = "Hello World"
    Public Sub New()
    End Sub
    Public Function GetGreeting() As String
        Return greeting
    End Function
End Class
Dim f As New Form1()
__Check(CStr(f.GetGreeting()), "Hello World")
