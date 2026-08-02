' vybe-test: vb/vb_spec_utility_classes/utility_spec_stopwatch_state_transitions
' origin: languages/vb/tests/vb/test_vb_spec_utility_classes.rs

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

Imports System.Diagnostics
Module Program
    Sub Main()
        Dim sw As New Stopwatch()
        sw.Start()
        __Check(CStr(sw.IsRunning), "True")
        sw.Stop()
        __Check(CStr(sw.IsRunning), "False")
        __Check(CStr(sw.ElapsedMilliseconds >= 0), "True")
        sw.Reset()
        __Check(CStr(sw.ElapsedMilliseconds), "0")
        sw.Restart()
        __Check(CStr(sw.IsRunning), "True")
    End Sub
End Module
