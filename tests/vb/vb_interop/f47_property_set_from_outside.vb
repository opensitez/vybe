' vybe-test: vb/vb_interop/f47_property_set_from_outside
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

Public Class Holder
    Dim data As String
    Public Sub New()
        data = ""
    End Sub
    Public Function GetData() As String
        Return data
    End Function
End Class
Dim h As New Holder()
h.data = "external"
__Check(CStr(h.GetData()), "external")
