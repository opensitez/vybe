' vybe-test: vb/vb_forms/f06_form_computed_value
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
    Dim width As Integer = 100
    Dim height As Integer = 50
    Public Sub New()
    End Sub
    Public Function Area() As Integer
        Return width * height
    End Function
End Class
Dim f As New Form1()
__Check(CStr(f.Area()), "5000")
