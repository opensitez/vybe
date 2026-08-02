' vybe-test: vb/vb_xml_linq_transformation_pipeline/test_vb_xml_save_and_load_memory_stream
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

Imports System.IO
Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim origDoc As New XDocument(New XElement("Root", "StreamTest"))
        Using ms As New MemoryStream()
            origDoc.Save(ms)
            ms.Position = 0
            Dim restoredDoc = XDocument.Load(ms)
            __Check(CStr(restoredDoc.Root.Value), "StreamTest")
        End Using
    End Sub
End Module
