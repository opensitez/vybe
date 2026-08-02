' vybe-test: vb/vb_interop/b67_cross_class_method_chain
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

Public Class A
    Public Function GetVal() As Integer
        Return 5
    End Function
End Class
Public Class B
    Dim a As A
    Public Sub New()
        a = New A()
    End Sub
    Public Function GetDouble() As Integer
        Return a.GetVal() * 2
    End Function
End Class
Public Class C
    Dim b As B
    Public Sub New()
        b = New B()
    End Sub
    Public Function GetTriple() As Integer
        Return b.GetDouble() + b.GetDouble() + b.GetDouble()
    End Function
End Class
Dim c As New C()
__Check(CStr(c.GetTriple()), "30")
