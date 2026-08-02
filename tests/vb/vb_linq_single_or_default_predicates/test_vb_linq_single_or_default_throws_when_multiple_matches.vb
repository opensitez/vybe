' vybe-test: vb/vb_linq_single_or_default_predicates/test_vb_linq_single_or_default_throws_when_multiple_matches
' origin: languages/vb/tests/vb/test_vb_linq_single_or_default_predicates.rs

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
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {10, 20}
        Try
            Dim x = numbers.SingleOrDefault()
        Catch ex As InvalidOperationException
            __Check(CStr("SingleOrDefault Multiple Matches Exception Caught"), "SingleOrDefault Multiple Matches Exception Caught")
        End Try
    End Sub
End Module
