' vybe-test: vb/vb_linq_group_by_multi_key/test_vb_linq_group_by_aggregations
' origin: languages/vb/tests/vb/test_vb_linq_group_by_multi_key.rs

Imports System.Linq

Module Program
    Sub Main()
        Dim sales = {
            New With {.Region = "North", .Amount = 100D},
            New With {.Region = "North", .Amount = 200D},
            New With {.Region = "South", .Amount = 150D}
        }

        Dim summary = From s In sales
                      Group s By s.Region Into Total = Sum(s.Amount), Average = Average(s.Amount), Count()

        For Each sum In summary
            Console.WriteLine(sum.Region & ": Total=" & sum.Total & ", Count=" & sum.Count)
        Next
    End Sub
End Module
