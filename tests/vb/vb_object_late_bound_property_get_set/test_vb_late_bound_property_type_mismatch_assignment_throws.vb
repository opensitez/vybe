' vybe-test: vb/vb_object_late_bound_property_get_set/test_vb_late_bound_property_type_mismatch_assignment_throws
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

Imports System

Module Program
    Class StrictHolder
        Public Property Count As Integer
    End Class

    Sub Main()
        Dim obj As Object = New StrictHolder()
        Try
            obj.Count = "NotAnIntegerString"
        Catch ex As Exception
            __Check(CStr("Invalid Cast in Late Bound Property Set Caught"), "Invalid Cast in Late Bound Property Set Caught")
        End Try
    End Sub
End Module
