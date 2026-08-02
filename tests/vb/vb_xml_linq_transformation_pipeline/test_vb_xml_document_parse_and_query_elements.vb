' vybe-test: vb/vb_xml_linq_transformation_pipeline/test_vb_xml_document_parse_and_query_elements
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

Imports System.Linq
Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim xmlStr = "<Catalog><Item Price='10'>A</Item><Item Price='20'>B</Item></Catalog>"
        Dim doc = XDocument.Parse(xmlStr)

        Dim items = doc.Root.Elements("Item").Select(Function(e) e.Value & ":" & e.Attribute("Price").Value)
        __Check(CStr(String.Join(",", items)), "A:10,B:20")
    End Sub
End Module
