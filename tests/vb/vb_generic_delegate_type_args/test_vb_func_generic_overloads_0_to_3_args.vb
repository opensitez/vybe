' vybe-test: vb/vb_generic_delegate_type_args/test_vb_func_generic_overloads_0_to_3_args
' origin: languages/vb/tests/vb/test_vb_generic_delegate_type_args.rs

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
        Dim f0 As Func(Of String) = Function() "F0"
        Dim f1 As Func(Of Integer, String) = Function(i) "F1_" & i
        Dim f2 As Func(Of Integer, Integer, String) = Function(i, j) "F2_" & (i + j)
        __Check(CStr(f0() & "|" & f1(10) & "|" & f2(3, 4)), "F0|F1_10|F2_7")
    End Sub
End Module
