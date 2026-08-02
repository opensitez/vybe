' vybe-test: vb/vb_class_static_shared_constructor/test_vb_shared_constructor_base_derived_execution_order
' origin: languages/vb/tests/vb/test_vb_class_static_shared_constructor.rs

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

Class ParentClass
    Public Shared Step1 As String
    Shared Sub New()
        Step1 = "ParentSharedInit"
        __Check(CStr(Step1), "ChildSharedInit")
    End Sub
End Class

Class ChildClass
    Inherits ParentClass
    Public Shared Step2 As String
    Shared Sub New()
        Step2 = "ChildSharedInit"
        __Check(CStr(Step2), "ParentSharedInit")
    End Sub
End Class

Module Program
    Sub Main()
        Dim c As New ChildClass()
    End Sub
End Module
