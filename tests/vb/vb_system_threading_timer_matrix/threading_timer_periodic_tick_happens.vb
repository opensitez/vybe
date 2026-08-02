' vybe-test: vb/vb_system_threading_timer_matrix/threading_timer_periodic_tick_happens
' origin: languages/vb/tests/vb/test_vb_system_threading_timer_matrix.rs

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
Imports System.Threading

Module M
    Private Shared _tickCount As Integer

    Sub Main()
        _tickCount = 0
        Dim done As New AutoResetEvent(False)

        Using timer As New Timer(New TimerCallback(AddressOf OnPeriodicTick), done, 0, 5)
            done.WaitOne(2000)
        End Using

        __Check(CStr(_tickCount > 0), "True")
        __Check(CStr(_tickCount >= 1), "True")
    End Sub

    Private Shared Sub OnPeriodicTick(state As Object)
        Interlocked.Increment(_tickCount)
        Dim gate As AutoResetEvent = CType(state, AutoResetEvent)
        gate.Set()
    End Sub
End Module
