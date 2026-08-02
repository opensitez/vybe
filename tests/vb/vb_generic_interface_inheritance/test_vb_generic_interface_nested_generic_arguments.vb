' vybe-test: vb/vb_generic_interface_inheritance/test_vb_generic_interface_nested_generic_arguments
' origin: languages/vb/tests/vb/test_vb_generic_interface_inheritance.rs

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

Imports System.Collections.Generic

Interface IBatchProcessor(Of TCollection As IEnumerable(Of String))
    Function ProcessBatch(batch As TCollection) As String
End Interface

Class ListBatchProcessor
    Implements IBatchProcessor(Of List(Of String))
    Public Function ProcessBatch(batch As List(Of String)) As String Implements IBatchProcessor(Of List(Of String)).ProcessBatch
        Return String.Join(",", batch)
    End Function
End Class

Module Program
    Sub Main()
        Dim p As IBatchProcessor(Of List(Of String)) = New ListBatchProcessor()
        Dim items As New List(Of String) From {"A", "B", "C"}
        __Check(CStr(p.ProcessBatch(items)), "A,B,C")
    End Sub
End Module
