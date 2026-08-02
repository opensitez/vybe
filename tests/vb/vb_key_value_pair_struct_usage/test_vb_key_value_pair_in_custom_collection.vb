' vybe-test: vb/vb_key_value_pair_struct_usage/test_vb_key_value_pair_in_custom_collection
' origin: languages/vb/tests/vb/test_vb_key_value_pair_struct_usage.rs

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

Class Cache(Of K, V)
    Private items As New List(Of KeyValuePair(Of K, V))()
    Public Sub Put(k As K, v As V)
        items.Add(New KeyValuePair(Of K, V)(k, v))
    End Sub
    Public Function GetFirst() As KeyValuePair(Of K, V)
        Return items(0)
    End Function
End Class

Module Program
    Sub Main()
        Dim c As New Cache(Of String, String)()
        c.Put("Token", "ABC123")
        Dim kv = c.GetFirst()
        __Check(CStr(kv.Key & "=" & kv.Value), "Token=ABC123")
    End Sub
End Module
