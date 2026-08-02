// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_index_of_finds_element_offset
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> span=stackalloc int[4]{10,20,30,40}; __Check((span.IndexOf(30)).ToString(), "2");
