' vybe-test: vb/vb_advanced_linq_xml/linq_skip_take_chain
' origin: languages/vb/tests/vb/test_vb_advanced_linq_xml.rs

Imports System.Linq

Module M
    Sub Main()
        Dim nums = {1, 2, 3, 4, 5, 6, 7}
        Dim query = From n In nums
                    Skip 2
                    Take 3
                    Select n
                    
        For Each n In query
            Console.WriteLine(n)
        Next
    End Sub
End Module
