' vybe-test: vb/vb_xml_linq_transformation_pipeline/test_vb_xml_transformation_to_domain_objects
' origin: languages/vb/tests/vb/test_vb_xml_linq_transformation_pipeline.rs

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
Imports System.Linq
Imports System.Xml.Linq

Class Book
    Public Property Title As String
    Public Property Author As String
End Class

Module Program
    Sub Main()
        Dim xmlStr = "<Library><Book><Title>VB Guide</Title><Author>Alice</Author></Book></Library>"
        Dim doc = XDocument.Parse(xmlStr)

        Dim books = doc.Root.Elements("Book").Select(Function(b) New Book With {
            .Title = b.Element("Title").Value,
            .Author = b.Element("Author").Value
        }).ToList()

        __Check(CStr(books(0).Title & " by " & books(0).Author), "VB Guide by Alice")
    End Sub
End Module
