// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_starts_with_matching_prefix
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> span=stackalloc int[3]{1,2,3}; System.ReadOnlySpan<int> prefix=stackalloc int[2]{1,2}; __Check((span.StartsWith(prefix)).ToString(), "True");
