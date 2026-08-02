// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_empty_property_matches_zero_length
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> span=stackalloc int[0]; __Check((System.Span<int>.Empty.Length).ToString(), "0");
