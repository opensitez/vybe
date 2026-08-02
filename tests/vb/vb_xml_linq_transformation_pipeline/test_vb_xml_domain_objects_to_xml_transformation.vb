' vybe-test: vb/vb_xml_linq_transformation_pipeline/test_vb_xml_domain_objects_to_xml_transformation
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

Class Item
    Public Property Code As String
    Public Property Qty As Integer
End Class

Module Program
    Sub Main()
        Dim items As New List(Of Item) From {
            New Item With {.Code = "I1", .Qty = 10},
            New Item With {.Code = "I2", .Qty = 20}
        }

        Dim root As New XElement("Inventory",
            From i In items Select New XElement("Item", New XAttribute("Code", i.Code), i.Qty)
        )

        __Check(CStr(root.Elements("Item").Count() & "|" & root.ToString().Contains("Code=""I1""")), "2|True")
    End Sub
End Module
