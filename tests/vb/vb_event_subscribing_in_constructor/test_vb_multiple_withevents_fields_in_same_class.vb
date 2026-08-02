' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_multiple_withevents_fields_in_same_class
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

Class EmitterA
    Public Event EventA As EventHandler
    Public Sub Fire()
        RaiseEvent EventA(Me, EventArgs.Empty)
    End Sub
End Class

Class EmitterB
    Public Event EventB As EventHandler
    Public Sub Fire()
        RaiseEvent EventB(Me, EventArgs.Empty)
    End Sub
End Class

Class CombinedListener
    Public WithEvents SourceA As EmitterA
    Public WithEvents SourceB As EmitterB

    Private Sub OnA(sender As Object, e As EventArgs) Handles SourceA.EventA
        __Check(CStr("A Handled"), "A Handled")
    End Sub

    Private Sub OnB(sender As Object, e As EventArgs) Handles SourceB.EventB
        __Check(CStr("B Handled"), "B Handled")
    End Sub
End Class

Module Program
    Sub Main()
        Dim l As New CombinedListener()
        l.SourceA = New EmitterA()
        l.SourceB = New EmitterB()
        l.SourceA.Fire()
        l.SourceB.Fire()
    End Sub
End Module
