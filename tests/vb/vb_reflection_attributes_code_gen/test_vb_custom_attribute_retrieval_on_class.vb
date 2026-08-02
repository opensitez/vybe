' vybe-test: vb/vb_reflection_attributes_code_gen/test_vb_custom_attribute_retrieval_on_class
' origin: languages/vb/tests/vb/test_vb_reflection_attributes_code_gen.rs

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

<AttributeUsage(AttributeTargets.Class)>
Class AuthorAttribute
    Inherits Attribute
    Public ReadOnly Name As String
    Public Sub New(authorName As String)
        Name = authorName
    End Sub
End Class

<Author("Alice")>
Class Document
End Class

Module Program
    Sub Main()
        Dim t = GetType(Document)
        Dim attr = CType(Attribute.GetCustomAttribute(t, GetType(AuthorAttribute)), AuthorAttribute)
        __Check(CStr(attr.Name), "Alice")
    End Sub
End Module
