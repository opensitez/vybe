' vybe-test: vb/vb_forms_advanced/e23_partial_constructor_in_designer
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

Partial Public Class Form1
    Dim title As String = "My Form"
    Public Sub New()
    End Sub
End Class
Partial Public Class Form1
    Public Function GetTitle() As String
        Return title
    End Function
End Class
Dim f As New Form1()
__Check(CStr(f.GetTitle()), "My Form")
