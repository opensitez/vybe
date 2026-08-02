// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_binary_search_finds_existing_value
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> span=stackalloc int[5]{1,3,5,7,9}; __Check((System.MemoryExtensions.BinarySearch(span,5)).ToString(), "2");
