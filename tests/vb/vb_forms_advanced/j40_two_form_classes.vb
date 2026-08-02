' vybe-test: vb/vb_forms_advanced/j40_two_form_classes
' origin: languages/vb/tests/vb/vb_forms_advanced_test.rs

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
    Dim title As String = "Form One"
    Public Sub New()
    End Sub
    Public Function GetTitle() As String
        Return title
    End Function
End Class
Public Class Form2
    Dim title As String = "Form Two"
    Public Sub New()
    End Sub
    Public Function GetTitle() As String
        Return title
    End Function
End Class
Dim f1 As New Form1()
Dim f2 As New Form2()
__Check(CStr(f1.GetTitle()), "Form One")
__Check(CStr(f2.GetTitle()), "Form Two")
