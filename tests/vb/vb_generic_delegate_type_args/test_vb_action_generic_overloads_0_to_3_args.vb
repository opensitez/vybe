' vybe-test: vb/vb_generic_delegate_type_args/test_vb_action_generic_overloads_0_to_3_args
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
        Dim a0 As Action = Sub() __Check(CStr("A0"), "A0")
        Dim a1 As Action(Of String) = Sub(s) __Check(CStr("A1_" & s), "A1_X")
        Dim a2 As Action(Of String, Integer) = Sub(s, i) __Check(CStr("A2_" & s & "_" & i), "A2_Y_99")
        a0()
        a1("X")
        a2("Y", 99)
    End Sub
End Module
