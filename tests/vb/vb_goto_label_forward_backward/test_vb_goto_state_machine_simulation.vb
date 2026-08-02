' vybe-test: vb/vb_goto_label_forward_backward/test_vb_goto_state_machine_simulation
' origin: languages/vb/tests/vb/test_vb_goto_label_forward_backward.rs

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

Module Program
    Sub Main()
        Dim state = 0
State0:
        __Check(CStr("State0"), "State0")
        state = 1
        GoTo State1
State1:
        __Check(CStr("State1"), "State1")
        state = 2
        If state = 2 Then GoTo StateEnd
StateEnd:
        __Check(CStr("StateEnd"), "StateEnd")
    End Sub
End Module
