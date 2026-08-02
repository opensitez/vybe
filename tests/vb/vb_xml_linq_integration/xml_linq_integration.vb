' vybe-test: vb/vb_xml_linq_integration/xml_linq_integration
' origin: languages/vb/tests/vb/test_vb_xml_linq_integration.rs

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

Imports System.Linq

Module M
    Sub Main()
        Dim items = {1, 2, 3}
        
        ' Embedded expression with LINQ inside XML literal
        Dim xml = <Root>
                      <%= From x In items Select <Item><%= x %></Item> %>
                  </Root>
                  
        __Check(CStr(xml.<Item>.Count()), "3")
    End Sub
End Module
