' vybe-test: vb/vb_interop/g56_textbox_properties
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

Imports System.Windows.Forms
Imports System.Drawing
Dim txt As New TextBox()
txt.Name = "txtInput"
txt.Location = New Point(15, 25)
txt.Size = New Size(180, 22)
__Check(CStr(txt.name), "txtInput")
__Check(CStr(txt.location.x), "15")
__Check(CStr(txt.location.y), "25")
__Check(CStr(txt.size.width), "180")
__Check(CStr(txt.size.height), "22")
