' vybe-test: vb/vb_xml_linq_transformation_pipeline/test_vb_xml_ancestors_and_parent_navigation
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
        Dim doc = XDocument.Parse("<GrandParent><Parent><Child>Leaf</Child></Parent></GrandParent>")
        Dim child = doc.Descendants("Child").First()
        Dim parentName = child.Parent.Name.LocalName
        Dim rootName = child.Ancestors().Last().Name.LocalName
        __Check(CStr(parentName & "|" & rootName), "Parent|GrandParent")
    End Sub
End Module
