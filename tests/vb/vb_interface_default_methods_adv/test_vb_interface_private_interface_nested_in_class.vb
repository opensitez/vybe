' vybe-test: vb/vb_interface_default_methods_adv/test_vb_interface_private_interface_nested_in_class
' origin: languages/vb/tests/vb/test_vb_interface_default_methods_adv.rs

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

Class Outer
    Private Interface IInternal
        Sub InternalWork()
    End Interface
    Private Class Inner
        Implements IInternal
        Public Sub InternalWork() Implements IInternal.InternalWork
            __Check(CStr("Internal Work Done"), "Internal Work Done")
        End Sub
    End Class
    Public Sub Run()
        Dim i As IInternal = New Inner()
        i.InternalWork()
    End Sub
End Class

Module Program
    Sub Main()
        Dim o As New Outer()
        o.Run()
    End Sub
End Module
