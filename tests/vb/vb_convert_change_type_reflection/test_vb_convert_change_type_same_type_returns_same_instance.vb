' vybe-test: vb/vb_convert_change_type_reflection/test_vb_convert_change_type_same_type_returns_same_instance
' origin: languages/vb/tests/vb/test_vb_convert_change_type_reflection.rs

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

Imports System

Module Program
    Sub Main()
        Dim orig As Object = "SameString"
        Dim converted As Object = Convert.ChangeType(orig, GetType(String))
        __Check(CStr(Object.ReferenceEquals(orig, converted)), "True")
    End Sub
End Module
