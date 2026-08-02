' vybe-test: vb/vb_advanced_linq_xml/linq_let_multiple_variables
' origin: languages/vb/tests/vb/test_vb_advanced_linq_xml.rs

Imports System.Linq

Module M
    Sub Main()
        Dim strings = {"Apple", "Banana"}
        
        Dim query = From s In strings
                    Let first = s(0)
                    Let len = s.Length
                    Select first & len
                    
        For Each res In query
            Console.WriteLine(res)
        Next
    End Sub
End Module
