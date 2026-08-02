' vybe-test: vb/vb_forms_advanced/j41_form_holds_reference_to_other
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

Public Class Form2
    Dim msg As String = "from form2"
    Public Sub New()
    End Sub
    Public Function GetMsg() As String
        Return msg
    End Function
End Class
Public Class Form1
    Dim child As Form2
    Public Sub New()
        child = New Form2()
    End Sub
    Public Function GetChildMsg() As String
        Return child.GetMsg()
    End Function
End Class
Dim f As New Form1()
__Check(CStr(f.GetChildMsg()), "from form2")
