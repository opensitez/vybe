' vybe-test: vb/vb_version_class_comparison/test_vb_version_linq_query_filtering
' origin: languages/vb/tests/vb/test_vb_version_class_comparison.rs

Imports System
Imports System.Linq

Module Program
    Sub Main()
        Dim list = {New Version(1, 0), New Version(2, 0), New Version(3, 0)}
        Dim minV2 = list.Where(Function(v) v >= New Version(2, 0))
        For Each v In minV2
            Console.WriteLine(v.ToString())
        Next
    End Sub
End Module
