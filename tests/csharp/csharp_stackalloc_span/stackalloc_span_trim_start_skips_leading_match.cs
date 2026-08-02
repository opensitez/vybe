// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_trim_start_skips_leading_match
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> span=stackalloc int[4]{0,0,1,2}; var trimmed=span.TrimStart(0); __Check((trimmed[0]).ToString(), "1");
