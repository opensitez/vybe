' vybe-test: vb/vb_objects_collections/f51_chain_functions_pass_object
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

Class Data
    Public Value As String
End Class
Function CreateData() As Data
    Dim d As New Data()
    d.Value = "start"
    Return d
End Function
Sub TransformData(d As Data)
    d.Value = d.Value & "-transformed"
End Sub
Sub FinalizeData(d As Data)
    d.Value = d.Value & "-done"
End Sub
Dim d As Data = CreateData()
TransformData(d)
FinalizeData(d)
__Check(CStr(d.Value), "start-transformed-done")
