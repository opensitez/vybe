' vybe-test: vb/vb_xml_adv_cdata_ns/xml_literals_cdata
' origin: languages/vb/tests/vb/test_vb_xml_adv_cdata_ns.rs

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
        Dim xml = <Data><![CDATA[<Test> & "Quotes"]]></Data>
        __Check(CStr(xml.Value), "<Test> & ""Quotes""")
    End Sub
End Module
