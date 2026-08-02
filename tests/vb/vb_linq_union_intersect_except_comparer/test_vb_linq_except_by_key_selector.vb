' vybe-test: vb/vb_linq_union_intersect_except_comparer/test_vb_linq_except_by_key_selector
' origin: languages/vb/tests/vb/test_vb_linq_union_intersect_except_comparer.rs

Imports System.Linq

Class Product
    Public Property SKU As String
    Public Sub New(s As String) : SKU = s : End Sub
End Class

Module Program
    Sub Main()
        Dim prods = {New Product("A101"), New Product("B202"), New Product("C303")}
        Dim excludedSkus = {"B202"}
        Dim res = prods.ExceptBy(excludedSkus, Function(p) p.SKU)
        For Each p In res
            Console.WriteLine(p.SKU)
        Next
    End Sub
End Module
