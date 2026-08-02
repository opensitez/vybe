' vybe-test: vb/vb_interop/f48_method_returns_object_access_properties
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

Public Class Pair
    Dim first As String
    Dim second As String
    Public Sub New(a As String, b As String)
        first = a
        second = b
    End Sub
End Class
Public Class Factory
    Public Function MakePair() As Pair
        Return New Pair("hello", "world")
    End Function
End Class
Dim f As New Factory()
Dim p As Pair = f.MakePair()
__Check(CStr(p.first), "hello")
__Check(CStr(p.second), "world")
