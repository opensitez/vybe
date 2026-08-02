' vybe-test: vb/vb_is_operator_runtime_type_check/test_vb_typeof_is_nullable_value_type
' origin: languages/vb/tests/vb/test_vb_is_operator_runtime_type_check.rs

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
        Dim n1 As Object = CType(100, Integer?)
        Dim n2 As Object = CType(Nothing, Integer?)
        __Check(CStr(TypeOf n1 Is Integer & "|" & (n2 Is Nothing)), "True|True")
    End Sub
End Module
