' vybe-test: vb/vb_xml_literals/xml_literal_embedded_expressions
' origin: languages/vb/tests/vb/test_vb_xml_literals.rs

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
        Dim year As Integer = 2026
        Dim xml = <report year=<%= year %>>
                      <status>Complete</status>
                  </report>
                  
        __Check(CStr(xml.@year), "2026")
        __Check(CStr(xml.<status>.Value), "Complete")
    End Sub
End Module
