' vybe-test: vb/vb_linq_of_type_filtering/test_vb_linq_of_type_class_hierarchy
' origin: languages/vb/tests/vb/test_vb_linq_of_type_filtering.rs

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

Imports System.Linq

Class Base
End Class

Class ChildA
    Inherits Base
End Class

Class ChildB
    Inherits Base
End Class

Module Program
    Sub Main()
        Dim list As Base() = {New ChildA(), New ChildB(), New ChildA()}
        Dim onlyA = list.OfType(Of ChildA)()
        __Check(CStr(onlyA.Count()), "2")
    End Sub
End Module
