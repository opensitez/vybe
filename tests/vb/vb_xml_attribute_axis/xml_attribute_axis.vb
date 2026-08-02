' vybe-test: vb/vb_xml_attribute_axis/xml_attribute_axis
' origin: languages/vb/tests/vb/test_vb_xml_attribute_axis.rs

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
        Dim xml = <Data id="42" name="Item" />
                  
        ' XML attribute axis operator @
        __Check(CStr(xml.@id), "42")
        __Check(CStr(xml.@name), "Item")
    End Sub
End Module
