' vybe-test: vb/vb_xml_axis_properties/xml_axis_properties_attributes
' origin: languages/vb/tests/vb/test_vb_xml_axis_properties.rs

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
        ' XML literals
        Dim book = <book id="123"><title>Programming in VB.NET</title><author>Jane Doe</author></book>
        
        ' Child axis property
        __Check(CStr(book.<title>.Value), "Programming in VB.NET")
        
        ' Attribute axis property
        __Check(CStr(book.@id), "123")
        
        ' XML interpolation
        Dim newTitle = "VB.NET Advanced"
        Dim book2 = <book><title><%= newTitle %></title></book>
        __Check(CStr(book2.<title>.Value), "VB.NET Advanced")
    End Sub
End Module
