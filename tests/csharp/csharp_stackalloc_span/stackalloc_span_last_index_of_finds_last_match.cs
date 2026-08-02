// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_last_index_of_finds_last_match
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> span=stackalloc int[4]{1,2,2,3}; __Check((span.LastIndexOf(2)).ToString(), "2");
