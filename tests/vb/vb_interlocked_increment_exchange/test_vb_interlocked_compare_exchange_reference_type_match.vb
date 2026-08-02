' vybe-test: vb/vb_interlocked_increment_exchange/test_vb_interlocked_compare_exchange_reference_type_match
' origin: languages/vb/tests/vb/test_vb_interlocked_increment_exchange.rs

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

Imports System.Threading

Class Document
    Public Title As String
    Public Sub New(t As String) : Title = t : End Sub
End Class

Module Program
    Sub Main()
        Dim doc1 As New Document("D1")
        Dim doc2 As New Document("D2")
        Dim target As Document = doc1

        Dim oldVal = Interlocked.CompareExchange(target, doc2, doc1)
        __Check(CStr(Object.ReferenceEquals(oldVal, doc1) & "|" & target.Title), "True|D2")
    End Sub
End Module
