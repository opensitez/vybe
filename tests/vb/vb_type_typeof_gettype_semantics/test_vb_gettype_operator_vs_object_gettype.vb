' vybe-test: vb/vb_type_typeof_gettype_semantics/test_vb_gettype_operator_vs_object_gettype
' origin: languages/vb/tests/vb/test_vb_type_typeof_gettype_semantics.rs

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

Class Animal
End Class

Class Dog
    Inherits Animal
End Class

Module Program
    Sub Main()
        Dim d As Animal = New Dog()

        Dim staticType As Type = GetType(Animal)
        Dim runtimeType As Type = d.GetType()

        __Check(CStr(staticType.Name), "Animal")
        __Check(CStr(runtimeType.Name), "Dog")
        __Check(CStr(staticType = runtimeType), "False")
    End Sub
End Module
