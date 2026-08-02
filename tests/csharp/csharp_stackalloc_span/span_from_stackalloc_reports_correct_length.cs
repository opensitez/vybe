// vybe-test: csharp/csharp_stackalloc_span/span_from_stackalloc_reports_correct_length
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> span=stackalloc int[5]{1,2,3,4,5}; __Check((span.Length).ToString(), "5");
