' vybe-test: vb/vb_array_find_findindex_predicates/test_vb_array_find_enum_values
' origin: languages/vb/tests/vb/test_vb_array_find_findindex_predicates.rs

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
    Pending
    Active
    Inactive
End Enum

Module Program
    Sub Main()
        Dim list As Status() = {Status.Pending, Status.Active, Status.Inactive}
        Dim match As Status = Array.Find(list, Function(s) s = Status.Active)
        __Check(CStr(match.ToString()), "Active")
    End Sub
End Module
