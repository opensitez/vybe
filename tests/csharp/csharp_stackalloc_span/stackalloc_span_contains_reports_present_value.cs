// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_contains_reports_present_value
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> span=stackalloc int[3]{4,5,6}; __Check((span.Contains(5)).ToString(), "True");
