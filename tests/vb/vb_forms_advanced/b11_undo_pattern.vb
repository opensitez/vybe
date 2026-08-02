' vybe-test: vb/vb_forms_advanced/b11_undo_pattern
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
    Dim currentValue As String = "initial"
    Dim previousValue As String = ""
    Public Sub New()
    End Sub
    Public Sub SetValue(val As String)
        previousValue = currentValue
        currentValue = val
    End Sub
    Public Sub Undo()
        currentValue = previousValue
    End Sub
    Public Function GetValue() As String
        Return currentValue
    End Function
End Class
Dim f As New Form1()
f.SetValue("second")
f.SetValue("third")
__Check(CStr(f.GetValue()), "third")
f.Undo()
__Check(CStr(f.GetValue()), "second")
