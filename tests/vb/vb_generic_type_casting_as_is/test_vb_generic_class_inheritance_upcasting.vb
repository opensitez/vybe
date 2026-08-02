' vybe-test: vb/vb_generic_type_casting_as_is/test_vb_generic_class_inheritance_upcasting
' origin: languages/vb/tests/vb/test_vb_generic_type_casting_as_is.rs

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

Class Animal : End Class
Class Dog : Inherits Animal : End Class

Module Program
    Private Function Upcast(Of TDerived As TBase, TBase As Class)(item As TDerived) As TBase
        Return DirectCast(CObj(item), TBase)
    End Function

    Sub Main()
        Dim d As New Dog()
        Dim a As Animal = Upcast(Of Dog, Animal)(d)
        __Check(CStr(a IsNot Nothing), "True")
    End Sub
End Module
