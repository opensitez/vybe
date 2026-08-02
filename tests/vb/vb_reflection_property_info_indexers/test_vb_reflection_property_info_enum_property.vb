' vybe-test: vb/vb_reflection_property_info_indexers/test_vb_reflection_property_info_enum_property
' origin: languages/vb/tests/vb/test_vb_reflection_property_info_indexers.rs

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

Enum Level
    Low
    High
End Enum

Class TaskItem
    Public Property Priority As Level = Level.High
End Class

Module Program
    Sub Main()
        Dim item As New TaskItem()
        Dim prop = GetType(TaskItem).GetProperty("Priority")
        Dim val As Level = CType(prop.GetValue(item), Level)
        __Check(CStr(val.ToString()), "High")
    End Sub
End Module
