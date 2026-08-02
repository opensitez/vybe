' vybe-test: vb/vb_generic_constraints_multiple/test_vb_generic_constraint_structure_value_type
' origin: languages/vb/tests/vb/test_vb_generic_constraints_multiple.rs

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

Class ValueHolder(Of T As Structure)
    Public Value As T
    Public Sub New(v As T)
        Me.Value = v
    End Sub
End Class

Module Program
    Sub Main()
        Dim h As New ValueHolder(Of Integer)(42)
        __Check(CStr(h.Value), "42")
    End Sub
End Module
