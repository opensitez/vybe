' vybe-test: vb/vb_system_converter_matrix/converter_change_type_from_string
' origin: languages/vb/tests/vb/test_vb_system_converter_matrix.rs

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

Module M
    Sub Main()
        Dim value As Object = Convert.ChangeType("17", GetType(Integer))
        Dim asInt As Integer = CType(value, Integer)
        __Check(CStr(asInt), "17")
    End Sub
End Module
