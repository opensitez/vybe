' vybe-test: vb/vb_xml_linq_transformation_pipeline/test_vb_xml_namespaces_xnamespace_usage
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

Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim ns As XNamespace = "http://example.com/ns"
        Dim elem As New XElement(ns + "Root",
            New XAttribute(XNamespace.Xmlns + "ex", ns),
            New XElement(ns + "Child", "Value")
        )
        __Check(CStr(elem.Element(ns + "Child").Value), "Value")
    End Sub
End Module
