' vybe-test: vb/vb_interface_hierarchy_multiple_inheritance/test_vb_interface_trycast_safe_conversion
' origin: languages/vb/tests/vb/test_vb_interface_hierarchy_multiple_inheritance.rs

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

Interface IDisposableResource
    Sub CleanUp()
End Interface

Class SafeResource
    Implements IDisposableResource
    Public Sub CleanUp() Implements IDisposableResource.CleanUp
        __Check(CStr("Cleaned Up"), "Cleaned Up")
    End Sub
End Class

Module Program
    Sub Main()
        Dim obj As Object = New SafeResource()
        Dim res = TryCast(obj, IDisposableResource)
        If res IsNot Nothing Then
            res.CleanUp()
        End If
    End Sub
End Module
