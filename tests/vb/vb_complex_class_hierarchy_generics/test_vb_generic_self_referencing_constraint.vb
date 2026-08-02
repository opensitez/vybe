' vybe-test: vb/vb_complex_class_hierarchy_generics/test_vb_generic_self_referencing_constraint
' origin: languages/vb/tests/vb/test_vb_complex_class_hierarchy_generics.rs

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

Class ComparableBase(Of T As ComparableBase(Of T))
    Implements IComparable(Of T)
    Public Property Priority As Integer

    Public Function CompareTo(other As T) As Integer Implements IComparable(Of T).CompareTo
        Return Priority.CompareTo(other.Priority)
    End Function
End Class

Class DerivedItem
    Inherits ComparableBase(Of DerivedItem)
End Class

Module Program
    Sub Main()
        Dim item1 As New DerivedItem With {.Priority = 10}
        Dim item2 As New DerivedItem With {.Priority = 20}
        __Check(CStr(item1.CompareTo(item2)), "-1")
    End Sub
End Module
