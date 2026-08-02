// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_index_from_end_reads_last_slot
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> span=stackalloc int[3]{8,9,10}; __Check((span[^1]).ToString(), "10");
