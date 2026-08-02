' vybe-test: vb/vb_objects_collections/f52_object_in_collection_method_called
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

Class Calc
    Public Base As Integer
    Public Function Double() As Integer
        Return Base * 2
    End Function
End Class
Dim list As New List(Of Calc)
Dim c As New Calc()
c.Base = 21
list.Add(c)
Dim retrieved As Calc = list.Item(0)
__Check(CStr(retrieved.Double()), "42")
