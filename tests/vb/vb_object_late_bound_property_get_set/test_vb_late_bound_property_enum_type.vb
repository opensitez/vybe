' vybe-test: vb/vb_object_late_bound_property_get_set/test_vb_late_bound_property_enum_type
' origin: languages/vb/tests/vb/test_vb_object_late_bound_property_get_set.rs

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

Enum PriorityLevel
    Low = 1
    Critical = 10
End Enum

Module Program
    Class TaskItem
        Public Property Priority As PriorityLevel
    End Class

    Sub Main()
        Dim obj As Object = New TaskItem()
        obj.Priority = PriorityLevel.Critical
        Dim p As PriorityLevel = CType(obj.Priority, PriorityLevel)
        __Check(CStr(p.ToString()), "Critical")
    End Sub
End Module
