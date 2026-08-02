' vybe-test: vb/vb_custom_event_thread_sync/test_vb_custom_event_async_event_raising
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
Imports System.Threading.Tasks

Class AsyncEventSource
    Public Event AsyncNotice As EventHandler

    Public Async Function FireAsync() As Task
        Await Task.Yield()
        RaiseEvent AsyncNotice(Me, EventArgs.Empty)
    End Function
End Class

Module Program
    Sub Main()
        Dim src As New AsyncEventSource()
        AddHandler src.AsyncNotice, Sub(s, e) __Check(CStr("Async Notice Fired"), "Async Notice Fired")
        Dim t = src.FireAsync()
        t.Wait()
    End Sub
End Module
