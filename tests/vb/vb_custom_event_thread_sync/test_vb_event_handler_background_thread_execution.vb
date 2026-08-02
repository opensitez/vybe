' vybe-test: vb/vb_custom_event_thread_sync/test_vb_event_handler_background_thread_execution
' origin: languages/vb/tests/vb/test_vb_custom_event_thread_sync.rs

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

Class BackgroundWorkerNotifier
    Public Event WorkDone As EventHandler

    Public Sub StartBackgroundJob()
        Dim t As New Thread(Sub()
            Thread.Sleep(10)
            RaiseEvent WorkDone(Me, EventArgs.Empty)
        End Sub)
        t.Start()
        t.Join()
    End Sub
End Class

Module Program
    Sub Main()
        Dim bwn As New BackgroundWorkerNotifier()
        AddHandler bwn.WorkDone, Sub(s, e) __Check(CStr("Done on Thread: " & Thread.CurrentThread.IsBackground), "Done on Thread: True")
        bwn.StartBackgroundJob()
    End Sub
End Module
