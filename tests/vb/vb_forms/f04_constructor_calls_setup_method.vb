' vybe-test: vb/vb_forms/f04_constructor_calls_setup_method
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
    Dim title As String
    Dim w As Integer
    Dim h As Integer
    Public Sub New()
        SetupDefaults()
    End Sub
    Private Sub SetupDefaults()
        title = "Default"
        w = 800
        h = 600
    End Sub
    Public Function GetTitle() As String
        Return title
    End Function
    Public Function GetWidth() As Integer
        Return w
    End Function
End Class
Dim f As New Form1()
__Check(CStr(f.GetTitle()), "Default")
__Check(CStr(f.GetWidth()), "800")
