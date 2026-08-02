' vybe-test: vb/vb_advanced_linq_xml/linq_from_multiple
' origin: languages/vb/tests/vb/test_vb_advanced_linq_xml.rs

Imports System.Linq

Module M
    Sub Main()
        Dim arr1 = {1, 2}
        Dim arr2 = {"A", "B"}
        
        Dim query = From a In arr1, b In arr2
                    Select a & b
                    
        For Each q In query
            Console.WriteLine(q)
        Next
    End Sub
End Module
