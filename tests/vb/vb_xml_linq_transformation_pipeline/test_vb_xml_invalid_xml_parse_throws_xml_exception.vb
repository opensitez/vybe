' vybe-test: vb/vb_xml_linq_transformation_pipeline/test_vb_xml_invalid_xml_parse_throws_xml_exception
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

Imports System.Xml
Imports System.Xml.Linq

Module Program
    Sub Main()
        Try
            XDocument.Parse("<UnclosedTag>Content")
        Catch ex As XmlException
            __Check(CStr("XmlException Caught on Malformed XML"), "XmlException Caught on Malformed XML")
        End Try
    End Sub
End Module
