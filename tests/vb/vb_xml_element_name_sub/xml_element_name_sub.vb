' vybe-test: vb/vb_xml_element_name_sub/xml_element_name_sub
' origin: languages/vb/tests/vb/test_vb_xml_element_name_sub.rs

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
        Dim name = "DynamicName"
        ' Element name substitution
        Dim xml = <<%= name %>>Content</>
                  
        __Check(CStr(xml.Name.LocalName), "DynamicName")
    End Sub
End Module
