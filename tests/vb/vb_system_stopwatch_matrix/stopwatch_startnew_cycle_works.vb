' vybe-test: vb/vb_system_stopwatch_matrix/stopwatch_startnew_cycle_works
' origin: languages/vb/tests/vb/test_vb_system_stopwatch_matrix.rs

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
Imports System.Diagnostics

Module M
    Sub Main()
        Dim sw As Stopwatch = Stopwatch.StartNew()
        sw.Stop()
        __Check(CStr(sw.IsRunning), "False")
        __Check(CStr(sw.ElapsedMilliseconds >= 0), "True")
        sw.Reset()
        __Check(CStr(sw.ElapsedMilliseconds), "0")
        __Check(CStr(sw.ElapsedTicks), "0")
    End Sub
End Module
