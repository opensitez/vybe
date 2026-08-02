' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_withevents_structure_not_supported_uses_class
' origin: languages/vb/tests/vb/test_vb_event_subscribing_in_constructor.rs

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

Class StructEventSource
    Public Event Signal As EventHandler
    Public Sub Fire()
        RaiseEvent Signal(Me, EventArgs.Empty)
    End Sub
End Class

Class Controller
    Public WithEvents Source As StructEventSource

    Private Sub OnSignal(sender As Object, e As EventArgs) Handles Source.Signal
        __Check(CStr("Signal Processed"), "Signal Processed")
    End Sub
End Class

Module Program
    Sub Main()
        Dim c As New Controller With {.Source = New StructEventSource()}
        c.Source.Fire()
    End Sub
End Module
