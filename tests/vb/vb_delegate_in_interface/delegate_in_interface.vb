' vybe-test: vb/vb_delegate_in_interface/delegate_in_interface
' origin: languages/vb/tests/vb/test_vb_delegate_in_interface.rs

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

' VB.NET allows defining delegates inside namespaces, modules, or classes, but in Interfaces?
' Interfaces cannot contain nested types in C#, but in VB.NET? Actually VB doesn't allow nested types in interfaces.
' Let's define a delegate outside and an event of that delegate inside.
Delegate Sub MyCustomHandler(msg As String)

Interface INotifier
    Event Notified As MyCustomHandler
End Interface

Class Notifier
    Implements INotifier
    Public Event Notified As MyCustomHandler Implements INotifier.Notified
    
    Public Sub Trigger()
        RaiseEvent Notified("Alert")
    End Sub
End Class

Module M
    Sub Handle(msg As String)
        __Check(CStr(msg), "Alert")
    End Sub

    Sub Main()
        Dim n As New Notifier()
        AddHandler n.Notified, AddressOf Handle
        n.Trigger()
    End Sub
End Module
