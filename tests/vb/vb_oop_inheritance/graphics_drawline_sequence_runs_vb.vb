' vybe-test: vb/vb_oop_inheritance/graphics_drawline_sequence_runs_vb
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

Imports System.Drawing
Imports System.Windows.Forms
Module M
Sub Main()
Dim g As Graphics = New PictureBox().CreateGraphics()
Dim p As New Pen(Color.Red, 2)
g.DrawLine(p, 0, 0, 10, 10)
__Check(CStr("drew"), "drew")
End Sub
End Module
