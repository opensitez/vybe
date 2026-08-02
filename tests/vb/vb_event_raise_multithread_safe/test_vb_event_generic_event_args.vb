' vybe-test: vb/vb_event_raise_multithread_safe/test_vb_event_generic_event_args
' origin: languages/vb/tests/vb/test_vb_event_raise_multithread_safe.rs

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

Class DataEventArgs(Of T)
    Inherits EventArgs
    Public ReadOnly Property Data As T
    Public Sub New(d As T)
        Data = d
    End Sub
End Class

Class DataBroadcaster(Of T)
    Public Event DataReceived As EventHandler(Of DataEventArgs(Of T))
    Public Sub Broadcast(d As T)
        RaiseEvent DataReceived(Me, New DataEventArgs(Of T)(d))
    End Sub
End Class

Module Program
    Sub Main()
        Dim b As New DataBroadcaster(Of String)()
        AddHandler b.DataReceived, Sub(s, e) __Check(CStr("Recv: " & e.Data), "Recv: PayloadString")
        b.Broadcast("PayloadString")
    End Sub
End Module
