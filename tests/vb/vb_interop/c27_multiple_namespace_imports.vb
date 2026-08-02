' vybe-test: vb/vb_interop/c27_multiple_namespace_imports
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

Imports System.Drawing
Imports System.Windows.Forms
Dim pt As New Point(5, 10)
Dim btn As New Button()
btn.Location = pt
__Check(CStr(btn.location.x), "5")
__Check(CStr(btn.location.y), "10")
