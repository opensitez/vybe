' vybe-test: vb/vb_complex_class_hierarchy_generics/test_vb_generic_enum_constraint_simulation
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

Module Program
    Private Function GetEnumName(Of T As {Structure, System.IConvertible})(val As T) As String
        Return [Enum].GetName(GetType(T), val)
    End Function

    Enum Status
        Active = 1
    End Enum

    Sub Main()
        Dim name = GetEnumName(Status.Active)
        __Check(CStr(name), "Active")
    End Sub
End Module
