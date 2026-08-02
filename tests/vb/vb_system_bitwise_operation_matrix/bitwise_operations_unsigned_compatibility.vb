' vybe-test: vb/vb_system_bitwise_operation_matrix/bitwise_operations_unsigned_compatibility
' origin: languages/vb/tests/vb/test_vb_system_bitwise_operation_matrix.rs

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
        Dim left As UInteger = 1UI
        Dim right As UInteger = &H80000000UI

        Dim combined As UInteger = left Or right
        __Check(CStr(combined = CUInt(&H80000001)), "True")
        __Check(CStr((combined And right) = right), "True")
    End Sub
End Module
