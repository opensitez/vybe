' vybe-test: vb/vb_array_true_for_all_exists/test_vb_array_trueforall_struct_array
' origin: languages/vb/tests/vb/test_vb_array_true_for_all_exists.rs

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

Structure Item
    Public ID As Integer
    Public Sub New(id As Integer)
        Me.ID = id
    End Sub
End Structure

Module Program
    Sub Main()
        Dim items As Item() = {New Item(1), New Item(2), New Item(3)}
        Dim allValid As Boolean = Array.TrueForAll(items, Function(i) i.ID > 0)
        __Check(CStr(allValid), "True")
    End Sub
End Module
