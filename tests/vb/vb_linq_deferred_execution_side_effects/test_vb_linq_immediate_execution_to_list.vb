' vybe-test: vb/vb_linq_deferred_execution_side_effects/test_vb_linq_immediate_execution_to_list
' origin: languages/vb/tests/vb/test_vb_linq_deferred_execution_side_effects.rs

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

Imports System.Collections.Generic
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers As New List(Of Integer) From {1, 2, 3}
        Dim snapshot = (From n In numbers Where n > 1 Select n * 10).ToList()

        numbers.Add(4) ' Source mutation does NOT affect materialized list

        __Check(CStr(String.Join(",", snapshot)), "20,30")
    End Sub
End Module
