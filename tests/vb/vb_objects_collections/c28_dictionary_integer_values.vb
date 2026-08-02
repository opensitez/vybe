' vybe-test: vb/vb_objects_collections/c28_dictionary_integer_values
' origin: languages/vb/tests/vb/vb_objects_collections_test.rs

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

Dim dict As New Dictionary(Of String, Integer)
dict.Add("score", 100)
dict.Add("bonus", 50)
Dim s As Integer = dict.Item("score") + dict.Item("bonus")
__Check(CStr(s), "150")
