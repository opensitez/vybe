' vybe-test: vb/vb_idisposable_double_dispose_safe/test_vb_using_statement_struct_disposable
' origin: languages/vb/tests/vb/test_vb_idisposable_double_dispose_safe.rs

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

Structure StructDisposable
    Implements IDisposable

    Public Property ID As Integer
    Public Sub New(i As Integer)
        ID = i
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        __Check(CStr("Struct Disposed " & ID), "Inside Struct Using")
    End Sub
End Structure

Module Program
    Sub Main()
        Using s As New StructDisposable(42)
            __Check(CStr("Inside Struct Using"), "Struct Disposed 42")
        End Using
    End Sub
End Module
