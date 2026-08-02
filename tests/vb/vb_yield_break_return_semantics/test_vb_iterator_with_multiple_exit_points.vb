' vybe-test: vb/vb_yield_break_return_semantics/test_vb_iterator_with_multiple_exit_points
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
    Private Iterator Function MultiExit(mode As Integer) As IEnumerable(Of String)
        Yield "Step1"
        If mode = 1 Then Return
        Yield "Step2"
        If mode = 2 Then Exit Function
        Yield "Step3"
    End Function

    Sub Main()
        __Check(CStr(String.Join("|", MultiExit(1))), "Step1")
        __Check(CStr(String.Join("|", MultiExit(2))), "Step1|Step2")
        __Check(CStr(String.Join("|", MultiExit(3))), "Step1|Step2|Step3")
    End Sub
End Module
