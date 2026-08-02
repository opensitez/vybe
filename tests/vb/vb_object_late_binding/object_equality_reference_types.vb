' vybe-test: vb/vb_object_late_binding/object_equality_reference_types
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
End Class
Module M
Sub Main()
Dim obj1 As Object = New C()
Dim obj2 As Object = New C()
Try
Dim r = (obj1 = obj2)
Catch
__Check(CStr("Caught"), "Caught")
End Try
End Sub
End Module
