' vybe-test: vb/vb_lazy_thread_safe_mode_execution/test_vb_lazy_list_collection_initialization
' origin: languages/vb/tests/vb/test_vb_lazy_thread_safe_mode_execution.rs

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
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim lazyList As New Lazy(Of List(Of String))(Function() New List(Of String) From {"A", "B", "C"})
        __Check(CStr(lazyList.Value.Count & ":" & String.Join(",", lazyList.Value)), "3:A,B,C")
    End Sub
End Module
