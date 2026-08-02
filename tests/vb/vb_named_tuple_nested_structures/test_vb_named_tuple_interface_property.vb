' vybe-test: vb/vb_named_tuple_nested_structures/test_vb_named_tuple_interface_property
' origin: languages/vb/tests/vb/test_vb_named_tuple_nested_structures.rs

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

Interface IScheduledTask
    ReadOnly Property ScheduleInfo As (StartTime As String, DurationMinutes As Integer)
End Interface

Class MaintenanceTask
    Implements IScheduledTask
    Public ReadOnly Property ScheduleInfo As (StartTime As String, DurationMinutes As Integer) Implements IScheduledTask.ScheduleInfo
        Get
            Return ("02:00", 60)
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim task As IScheduledTask = New MaintenanceTask()
        __Check(CStr(task.ScheduleInfo.StartTime & " for " & task.ScheduleInfo.DurationMinutes & "m"), "02:00 for 60m")
    End Sub
End Module
