' vybe-test: vb/vb_system_bitwise_operation_matrix/bitwise_operations_negative_and_not_behavior
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

Module M
    Sub Main()
        Dim x As Integer = -1
        Dim y As Integer = x And 3
        Dim z As Integer = x Or 0
        Dim n As Integer = Not 0

        __Check(CStr(y), "3")
        __Check(CStr(z), "-1")
        __Check(CStr(n = -1), "True")
    End Sub
End Module
