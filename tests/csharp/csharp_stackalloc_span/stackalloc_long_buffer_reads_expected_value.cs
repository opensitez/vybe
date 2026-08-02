// vybe-test: csharp/csharp_stackalloc_span/stackalloc_long_buffer_reads_expected_value
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<long> span=stackalloc long[2]{10000000000L,20000000000L}; __Check((span[0]>0).ToString(), "True");
