' vybe-test: vb/vb_event_handler_generic_args/test_vb_event_handler_generic_custom_event_args
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

Class OrderEventArgs
    Inherits EventArgs
    Public ReadOnly OrderId As Integer
    Public Sub New(id As Integer)
        Me.OrderId = id
    End Sub
End Class

Class OrderProcessor
    Public Event OrderProcessed As EventHandler(Of OrderEventArgs)

    Public Sub Process(id As Integer)
        RaiseEvent OrderProcessed(Me, New OrderEventArgs(id))
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New OrderProcessor()
        AddHandler p.OrderProcessed, Sub(sender, e)
            __Check(CStr("Order: " & e.OrderId), "Order: 1001")
        End Sub
        p.Process(1001)
    End Sub
End Module
