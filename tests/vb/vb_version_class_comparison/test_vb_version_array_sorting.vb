' vybe-test: vb/vb_version_class_comparison/test_vb_version_array_sorting
' origin: languages/vb/tests/vb/test_vb_version_class_comparison.rs

Imports System

Module Program
    Sub Main()
        Dim versions As Version() = {
            New Version(2, 0),
            New Version(1, 5),
            New Version(2, 0, 1),
            New Version(1, 0)
        }
        Array.Sort(versions)
        For Each v In versions
            Console.WriteLine(v.ToString())
        Next
    End Sub
End Module
