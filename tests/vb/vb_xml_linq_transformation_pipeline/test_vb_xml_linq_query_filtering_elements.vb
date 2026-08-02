' vybe-test: vb/vb_xml_linq_transformation_pipeline/test_vb_xml_linq_query_filtering_elements
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
        Dim doc As New XDocument(
            New XElement("Orders",
                New XElement("Order", New XAttribute("Status", "Completed"), "O1"),
                New XElement("Order", New XAttribute("Status", "Pending"), "O2"),
                New XElement("Order", New XAttribute("Status", "Completed"), "O3")
            )
        )

        Dim completed = From o In doc.Root.Elements("Order")
                        Where o.Attribute("Status").Value = "Completed"
                        Select o.Value

        __Check(CStr(String.Join(",", completed)), "O1,O3")
    End Sub
End Module
