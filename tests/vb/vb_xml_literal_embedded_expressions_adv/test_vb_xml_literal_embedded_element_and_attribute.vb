' vybe-test: vb/vb_xml_literal_embedded_expressions_adv/test_vb_xml_literal_embedded_element_and_attribute
' origin: languages/vb/tests/vb/test_vb_xml_literal_embedded_expressions_adv.rs

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
        Dim name As String = "Laptop"
        Dim price As Double = 999.99
        Dim id As Integer = 101

        Dim doc As XElement = <product id=<%= id %>>
                                  <name><%= name %></name>
                                  <price><%= price %></price>
                              </product>

        __Check(CStr(doc.@id), "101")
        __Check(CStr(doc.<name>.Value), "Laptop")
        __Check(CStr(doc.<price>.Value), "999.99")
    End Sub
End Module
