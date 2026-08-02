' vybe-test: vb/vb_xml_literals_expressions/xml_literals_expressions
' origin: languages/vb/tests/vb/test_vb_xml_literals_expressions.rs

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
        Dim name As String = "Bob"
        Dim age As Integer = 30
        
        ' Embedded expressions in XML literals use <%= expr %>
        Dim userXml As XElement = 
            <User>
                <Name><%= name %></Name>
                <Age><%= age %></Age>
            </User>
            
        __Check(CStr(userXml.<Name>.Value), "Bob")
        __Check(CStr(userXml.<Age>.Value), "30")
    End Sub
End Module
