' vybe-test: vb/vb_object_late_binding/late_binding_overload_resolution
' origin: languages/vb/tests/vb/test_vb_object_late_binding.rs

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

Option Strict Off
Class C
Public Function M(a As Integer) As String
Return "Int"
End Function
Public Function M(a As String) As String
Return "Str"
End Function
End Class
Module M
Sub Main()
Dim obj As Object = New C()
__Check(CStr(obj.M(10)), "Int")
End Sub
End Module
