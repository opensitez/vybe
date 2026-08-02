' vybe-test: vb/vb_handles_mybase/handles_mybase
' origin: languages/vb/tests/vb/test_vb_handles_mybase.rs

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

Class Base
    Public Event Processed As EventHandler
    
    Protected Sub Trigger()
        RaiseEvent Processed(Me, EventArgs.Empty)
    End Sub
End Class

Class Derived
    Inherits Base
    
    ' Handles MyBase.Event
    Private Sub OnProcessed(sender As Object, e As EventArgs) Handles MyBase.Processed
        __Check(CStr("Handled in Derived"), "Handled in Derived")
    End Sub
    
    Public Sub DoWork()
        Trigger()
    End Sub
End Class

Module M
    Sub Main()
        Dim d As New Derived()
        d.DoWork()
    End Sub
End Module
