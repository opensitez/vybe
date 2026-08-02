' vybe-test: vb/vb_generic_struct_methods/test_vb_generic_struct_constraint_icomparable
' origin: languages/vb/tests/vb/test_vb_generic_struct_methods.rs

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

Structure Range(Of T As IComparable(Of T))
    Public Min As T
    Public Max As T
    Public Sub New(min As T, max As T)
        Me.Min = min : Me.Max = max
    End Sub
    Public Function Contains(val As T) As Boolean
        Return val.CompareTo(Min) >= 0 AndAlso val.CompareTo(Max) <= 0
    End Function
End Structure

Module Program
    Sub Main()
        Dim r As New Range(Of Integer)(10, 20)
        __Check(CStr(r.Contains(15) & "|" & r.Contains(25)), "True|False")
    End Sub
End Module
