' vybe-test: vb/vb_generic_struct_constraints/test_vb_generic_struct_with_constraints
' origin: languages/vb/tests/vb/test_vb_generic_struct_constraints.rs

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

Structure Pair(Of TKey As IComparable, TValue)
    Public Key As TKey
    Public Value As TValue

    Public Sub New(k As TKey, v As TValue)
        Me.Key = k
        Me.Value = v
    End Sub
End Structure

Module Program
    Sub Main()
        Dim p As New Pair(Of Integer, String)(10, "Ten")
        __Check(CStr(p.Key & ":" & p.Value), "10:Ten")
    End Sub
End Module
