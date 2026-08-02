' vybe-test: vb/vb_array_true_for_all_exists/test_vb_array_exists_enum_array
' origin: languages/vb/tests/vb/test_vb_array_true_for_all_exists.rs

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

Enum Status
    Active
    Pending
    Inactive
End Enum

Module Program
    Sub Main()
        Dim states As Status() = {Status.Active, Status.Pending}
        Dim hasInactive As Boolean = Array.Exists(states, Function(s) s = Status.Inactive)
        __Check(CStr(hasInactive), "False")
    End Sub
End Module
