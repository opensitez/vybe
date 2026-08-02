' vybe-test: vb/vb_iif_ternary_evaluation/test_vb_if_operator_custom_class_inheritance_common_type
' origin: languages/vb/tests/vb/test_vb_iif_ternary_evaluation.rs

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

Class Animal
End Class

Class Dog
    Inherits Animal
End Class

Class Cat
    Inherits Animal
End Class

Module Program
    Sub Main()
        Dim isDog = True
        ' Common base class Animal is inferred!
        Dim pet As Animal = If(isDog, DirectCast(New Dog(), Animal), DirectCast(New Cat(), Animal))
        __Check(CStr(pet.GetType().Name), "Dog")
    End Sub
End Module
