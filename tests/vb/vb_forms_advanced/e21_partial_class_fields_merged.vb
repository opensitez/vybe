' vybe-test: vb/vb_forms_advanced/e21_partial_class_fields_merged
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
    Dim name As String = "hello"
End Class
Partial Public Class Form1
    Dim count As Integer = 42
    Public Sub New()
    End Sub
    Public Sub ShowBoth()
        __Check(CStr(name & " " & CStr(count)), "hello 42")
    End Sub
End Class
Dim f As New Form1()
f.ShowBoth()
