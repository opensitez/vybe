' vybe-test: vb/vb_xml_literals_namespaces/xml_literals_namespaces
' origin: languages/vb/tests/vb/test_vb_xml_literals_namespaces.rs

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

Imports <xmlns:ns="http://example.com/ns">

Module M
    Sub Main()
        Dim xml = <ns:Root>
                      <ns:Child>Value</ns:Child>
                  </ns:Root>
                  
        ' Need to use GetNamespace to query with namespaces
        Dim ns = GetXmlNamespace(ns)
        __Check(CStr(xml.Element(ns + "Child").Value), "Value")
    End Sub
End Module
