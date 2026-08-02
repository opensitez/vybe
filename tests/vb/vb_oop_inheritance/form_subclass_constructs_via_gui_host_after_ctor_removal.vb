' vybe-test: vb/vb_oop_inheritance/form_subclass_constructs_via_gui_host_after_ctor_removal
' origin: languages/vb/tests/vb/test_vb_oop_inheritance.rs

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

Imports System.Windows.Forms
Public Class MyForm
Inherits Form
End Class
Module M
Sub Main()
Dim f As New MyForm()
f.Text = "hello"
__Check(CStr(f.Text), "hello")
__Check(CStr(f.__control_type), "Form")
End Sub
End Module
