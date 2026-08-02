' vybe-test: vb/vb_advanced_linq_xml/xml_set_attribute_value
' origin: languages/vb/tests/vb/test_vb_advanced_linq_xml.rs

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

Module M
    Sub Main()
        Dim el = <Root Id="1"/>
        el.SetAttributeValue("Id", "2")
        el.SetAttributeValue("NewAttr", "3")
        
        __Check(CStr(el.Attribute("Id").Value), "2")
        __Check(CStr(el.Attribute("NewAttr").Value), "3")
    End Sub
End Module
