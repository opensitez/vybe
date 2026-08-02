' vybe-test: vb/vb_class_sealed_notinheritable_checks/test_vb_notinheritable_record_class_simulation
' origin: languages/vb/tests/vb/test_vb_class_sealed_notinheritable_checks.rs

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

NotInheritable Class ImmutableData
    Public ReadOnly Property ID As Integer
    Public ReadOnly Property Value As String
    Public Sub New(id As Integer, val As String)
        Me.ID = id : Me.Value = val
    End Sub
End Class

Module Program
    Sub Main()
        Dim d As New ImmutableData(1, "Val")
        __Check(CStr(d.ID & ":" & d.Value), "1:Val")
    End Sub
End Module
