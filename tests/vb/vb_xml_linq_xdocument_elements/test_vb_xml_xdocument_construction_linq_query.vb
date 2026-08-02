' vybe-test: vb/vb_xml_linq_xdocument_elements/test_vb_xml_xdocument_construction_linq_query
' origin: languages/vb/tests/vb/test_vb_xml_linq_xdocument_elements.rs

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

Imports System.Linq
Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim items = {"A", "B", "C"}
        Dim doc As New XDocument(
            New XElement("root",
                From i In items Select <item><%= i %></item>
            )
        )
        __Check(CStr(doc.Root.Elements("item").Count()), "3")
    End Sub
End Module
