' vybe-test: vb/vb_xml_literal_xmlns/xml_literal_xmlns
' origin: languages/vb/tests/vb/test_vb_xml_literal_xmlns.rs

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

Module M
    Sub Main()
        ' XML literal with inline xmlns
        Dim xml = <Root xmlns:ns="http://test.com">
                      <ns:Child>Val</ns:Child>
                  </Root>
                  
        __Check(CStr(xml.Name.LocalName), "Root")
    End Sub
End Module
