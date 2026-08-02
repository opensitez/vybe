' vybe-test: vb/vb_linq_all_any_predicates/test_vb_linq_all_enum_values_check
' origin: languages/vb/tests/vb/test_vb_linq_all_any_predicates.rs

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

Imports System.Linq

Enum Status
    Active
    Pending
End Enum

Module Program
    Sub Main()
        Dim statuses = {Status.Active, Status.Active}
        __Check(CStr(statuses.All(Function(s) s = Status.Active)), "True")
    End Sub
End Module
