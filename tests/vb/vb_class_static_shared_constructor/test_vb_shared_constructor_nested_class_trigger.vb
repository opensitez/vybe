' vybe-test: vb/vb_class_static_shared_constructor/test_vb_shared_constructor_nested_class_trigger
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

Class Parent
    Public Shared ParentInit As Boolean = False
    Shared Sub New()
        ParentInit = True
    End Sub

    Public Class Child
        Public Shared Function Work() As String
            Return "ChildWork"
        End Function
    End Class
End Class

Module Program
    Sub Main()
        Dim w = Parent.Child.Work()
        __Check(CStr(w & "|ParentInit=" & Parent.ParentInit), "ChildWork|ParentInit=True")
    End Sub
End Module
