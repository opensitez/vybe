' vybe-test: vb/vb_generic_static_member_per_type/test_vb_generic_shared_field_two_type_parameters
' origin: languages/vb/tests/vb/test_vb_generic_static_member_per_type.rs

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

Class MatrixTracker(Of T1, T2)
    Public Shared InstanceID As String
End Class

Module Program
    Sub Main()
        MatrixTracker(Of Integer, String).InstanceID = "IntString"
        MatrixTracker(Of Integer, Double).InstanceID = "IntDouble"

        __Check(CStr(MatrixTracker(Of Integer, String).InstanceID & "|" & MatrixTracker(Of Integer, Double).InstanceID), "IntString|IntDouble")
    End Sub
End Module
