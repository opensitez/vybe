// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_mismatch_reports_first_difference_index
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.ReadOnlySpan<int> a=stackalloc int[3]{1,2,3}; System.ReadOnlySpan<int> b=stackalloc int[3]{1,9,3}; __Check((a.Mismatch(b)).ToString(), "1");
