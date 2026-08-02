' vybe-test: vb/vb_action_func_delegates/test_vb_func_chaining_higher_order_function
' origin: languages/vb/tests/vb/test_vb_action_func_delegates.rs

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
    Function Combine(f1 As Func(Of Integer, Integer), f2 As Func(Of Integer, Integer)) As Func(Of Integer, Integer)
        Return Function(x) f2(f1(x))
    End Function

    Sub Main()
        Dim addTwo As Func(Of Integer, Integer) = Function(x) x + 2
        Dim mulThree As Func(Of Integer, Integer) = Function(x) x * 3
        Dim combined As Func(Of Integer, Integer) = Combine(addTwo, mulThree)
        __Check(CStr(combined(5)), "21")
    End Sub
End Module
