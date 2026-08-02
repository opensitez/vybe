' vybe-test: vb/vb_date_time_compare_is_leap_year/test_vb_date_time_array_sort
' origin: languages/vb/tests/vb/test_vb_date_time_compare_is_leap_year.rs

Imports System

Module Program
    Sub Main()
        Dim dates As DateTime() = {New DateTime(2025, 3, 1), New DateTime(2025, 1, 1), New DateTime(2025, 2, 1)}
        Array.Sort(dates)
        For Each d In dates
            Console.WriteLine(d.Month)
        Next
    End Sub
End Module
