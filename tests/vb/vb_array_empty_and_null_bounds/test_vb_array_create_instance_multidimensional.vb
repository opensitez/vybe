' vybe-test: vb/vb_array_empty_and_null_bounds/test_vb_array_create_instance_multidimensional
' origin: languages/vb/tests/vb/test_vb_array_empty_and_null_bounds.rs

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
        Dim lengths As Integer() = {2, 3}
        Dim grid As Array = Array.CreateInstance(GetType(Integer), lengths)
        grid.SetValue(42, 1, 2)
        __Check(CStr(grid.Rank), "2")
        __Check(CStr(grid.GetValue(1, 2)), "42")
    End Sub
End Module
