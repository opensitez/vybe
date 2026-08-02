' vybe-test: vb/vb_interop/f50_multiple_instances_independent_state
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

Public Class Counter
    Dim count As Integer
    Public Sub New(start As Integer)
        count = start
    End Sub
    Public Sub Inc()
        count = count + 1
    End Sub
    Public Function GetCount() As Integer
        Return count
    End Function
End Class
Dim a As New Counter(0)
Dim b As New Counter(100)
a.Inc()
a.Inc()
b.Inc()
__Check(CStr(a.GetCount()), "2")
__Check(CStr(b.GetCount()), "101")
