' vybe-test: vb/vb_system_timespan/system_timespan_arithmetic
' origin: languages/vb/tests/vb/test_vb_system_timespan.rs

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

Module M
    Sub Main()
        Dim ts1 As New TimeSpan(1, 0, 0)
        Dim ts2 As New TimeSpan(0, 30, 0)
        
        Dim ts3 = ts1 + ts2
        __Check(CStr(ts3.TotalMinutes), "90")
        
        Dim ts4 = ts1 - ts2
        __Check(CStr(ts4.TotalMinutes), "30")
    End Sub
End Module
