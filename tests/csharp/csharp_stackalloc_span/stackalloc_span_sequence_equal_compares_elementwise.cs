// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_sequence_equal_compares_elementwise
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> a=stackalloc int[2]{7,8}; System.Span<int> b=stackalloc int[2]{7,8}; __Check((a.SequenceEqual(b)).ToString(), "True");
