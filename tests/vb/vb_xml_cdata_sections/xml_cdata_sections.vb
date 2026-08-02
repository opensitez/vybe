' vybe-test: vb/vb_xml_cdata_sections/xml_cdata_sections
' origin: languages/vb/tests/vb/test_vb_xml_cdata_sections.rs

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
        ' XML literals support CDATA blocks
        Dim xml = <Data>
                      <![CDATA[Some <unescaped> data & characters]]>
                  </Data>
                  
        __Check(CStr(xml.Value.Trim()), "Some <unescaped> data & characters")
    End Sub
End Module
