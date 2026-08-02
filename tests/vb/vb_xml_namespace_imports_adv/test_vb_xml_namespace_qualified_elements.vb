' vybe-test: vb/vb_xml_namespace_imports_adv/test_vb_xml_namespace_qualified_elements
' origin: languages/vb/tests/vb/test_vb_xml_namespace_imports_adv.rs

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
Imports <xmlns:ns="http://example.com/ns">

Module Program
    Sub Main()
        Dim elem As XElement = <ns:data ns:attr="val">Content</ns:data>
        __Check(CStr(elem.Name.NamespaceName), "http://example.com/ns")
        __Check(CStr(elem.@ns:attr), "val")
    End Sub
End Module
