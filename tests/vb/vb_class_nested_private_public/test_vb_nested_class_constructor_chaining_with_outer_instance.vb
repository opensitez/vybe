' vybe-test: vb/vb_class_nested_private_public/test_vb_nested_class_constructor_chaining_with_outer_instance
' origin: languages/vb/tests/vb/test_vb_class_nested_private_public.rs

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

Class Document
    Public Title As String
    Public Sub New(t As String)
        Title = t
    End Sub

    Public Class Header
        Private parentDoc As Document
        Public Sub New(doc As Document)
            parentDoc = doc
        End Sub
        Public Function GetTitle() As String
            Return "Header of " & parentDoc.Title
        End Function
    End Class
End Class

Module Program
    Sub Main()
        Dim doc As New Document("Report")
        Dim h As New Document.Header(doc)
        __Check(CStr(h.GetTitle()), "Header of Report")
    End Sub
End Module
