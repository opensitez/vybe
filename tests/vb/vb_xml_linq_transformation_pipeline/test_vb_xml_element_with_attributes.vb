' vybe-test: vb/vb_xml_linq_transformation_pipeline/test_vb_xml_element_with_attributes
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
        Dim elem As New XElement("Product",
            New XAttribute("Id", "P100"),
            New XAttribute("Price", "29.99"),
            "Widget"
        )
        __Check(CStr(elem.Attribute("Id").Value & "|" & elem.Attribute("Price").Value & "|" & elem.Value), "P100|29.99|Widget")
    End Sub
End Module
