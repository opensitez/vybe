' vybe-test: vb/vb_yield_break_return_semantics/test_vb_yield_return_inside_select_case
' origin: languages/vb/tests/vb/test_vb_yield_break_return_semantics.rs

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

Module Program
    Private Iterator Function SelectGen(val As Integer) As IEnumerable(Of String)
        Select Case val
            Case 1
                Yield "One"
            Case 2
                Yield "TwoA"
                Yield "TwoB"
            Case Else
                Yield "Other"
        End Select
    End Function

    Sub Main()
        __Check(CStr(String.Join(",", SelectGen(2))), "TwoA,TwoB")
    End Sub
End Module
