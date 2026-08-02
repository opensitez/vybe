' vybe-test: vb/vb_interop/f51_property_get_set
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

Public Class Person
    Dim _name As String
    Public Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property
    Public Sub New()
        _name = "unknown"
    End Sub
End Class
Dim p As New Person()
__Check(CStr(p.Name), "unknown")
p.Name = "Alice"
__Check(CStr(p.Name), "Alice")
