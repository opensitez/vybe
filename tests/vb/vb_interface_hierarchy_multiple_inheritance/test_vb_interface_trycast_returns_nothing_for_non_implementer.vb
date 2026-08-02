' vybe-test: vb/vb_interface_hierarchy_multiple_inheritance/test_vb_interface_trycast_returns_nothing_for_non_implementer
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

Interface IRunnable
    Sub Run()
End Interface

Class NonRunnable
End Class

Module Program
    Sub Main()
        Dim obj As Object = New NonRunnable()
        Dim r = TryCast(obj, IRunnable)
        __Check(CStr(r Is Nothing), "True")
    End Sub
End Module
