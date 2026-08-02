' vybe-test: vb/vb_class_nested_private_public/test_vb_nested_class_factory_pattern
' origin: languages/vb/tests/vb/test_vb_class_nested_private_public.rs

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

Interface Product
    Function GetName() As String
End Interface

Class Factory
    Private Class ProductA
        Implements Product
        Public Function GetName() As String Implements Product.GetName : Return "ProdA" : End Function
    End Class

    Private Class ProductB
        Implements Product
        Public Function GetName() As String Implements Product.GetName : Return "ProdB" : End Function
    End Class

    Public Shared Function Create(type As String) As Product
        If type = "A" Then Return New ProductA()
        Return New ProductB()
    End Function
End Class

Module Program
    Sub Main()
        Dim p1 = Factory.Create("A")
        Dim p2 = Factory.Create("B")
        __Check(CStr(p1.GetName() & "&" & p2.GetName()), "ProdA&ProdB")
    End Sub
End Module
