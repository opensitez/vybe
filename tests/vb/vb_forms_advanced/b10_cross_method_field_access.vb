' vybe-test: vb/vb_forms_advanced/b10_cross_method_field_access
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
    Dim savedValue As String = ""
    Public Sub New()
    End Sub
    Public Sub Save(val As String)
        savedValue = val
    End Sub
    Public Function Load() As String
        Return savedValue
    End Function
End Class
Dim f As New Form1()
f.Save("important data")
__Check(CStr(f.Load()), "important data")
