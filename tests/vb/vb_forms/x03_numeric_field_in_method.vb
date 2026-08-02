' vybe-test: vb/vb_forms/x03_numeric_field_in_method
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

Public Class Counter
    Dim count As Integer = 0
    Public Sub New()
    End Sub
    Public Sub Add(n As Integer)
        count = count + n
    End Sub
    Public Function GetCount() As Integer
        Return count
    End Function
End Class
Dim c As New Counter()
c.Add(5)
c.Add(3)
__Check(CStr(c.GetCount()), "8")
