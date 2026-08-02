' vybe-test: vb/vb_event_handler_generic_args/test_vb_event_handler_cancel_event_args
' origin: languages/vb/tests/vb/test_vb_event_handler_generic_args.rs

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
Imports System.ComponentModel

Class Document
    Public Event Closing As EventHandler(Of CancelEventArgs)

    Public Function TryClose() As Boolean
        Dim args As New CancelEventArgs()
        RaiseEvent Closing(Me, args)
        Return Not args.Cancel
    End Function
End Class

Module Program
    Sub Main()
        Dim doc As New Document()
        AddHandler doc.Closing, Sub(sender, e)
            e.Cancel = True
        End Sub
        __Check(CStr(doc.TryClose()), "False")
    End Sub
End Module
