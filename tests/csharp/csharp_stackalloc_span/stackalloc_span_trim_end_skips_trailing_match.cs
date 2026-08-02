// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_trim_end_skips_trailing_match
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> span=stackalloc int[4]{1,2,0,0}; var trimmed=span.TrimEnd(0); __Check((trimmed[^1]).ToString(), "2");
