' vybe-test: vb/vb_interfaces_basic/interface_event_implementation
' origin: languages/vb/tests/vb/test_vb_interfaces_basic.rs

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

Interface IAlarm
    Event Triggered()
End Interface

Class SecuritySystem
    Implements IAlarm
    
    Public Event Triggered() Implements IAlarm.Triggered
    
    Public Sub SoundAlarm()
        RaiseEvent Triggered()
    End Sub
End Class

Module M
    Private WithEvents sys As SecuritySystem
    
    Private Sub sys_Triggered() Handles sys.Triggered
        __Check(CStr("Alert!"), "Alert!")
    End Sub
    
    Sub Main()
        sys = New SecuritySystem()
        sys.SoundAlarm()
    End Sub
End Module
