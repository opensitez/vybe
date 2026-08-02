' vybe-test: vb/vb_advanced_linq_xml/linq_let_multiple
' origin: languages/vb/tests/vb/test_vb_advanced_linq_xml.rs

Imports System.Linq

Module M
    Sub Main()
        Dim nums = {1, 2, 3}
        Dim query = From n In nums
                    Let x = n * 2, y = n * 3
                    Select x + y
                    
        For Each res In query
            Console.WriteLine(res)
        Next
    End Sub
End Module
