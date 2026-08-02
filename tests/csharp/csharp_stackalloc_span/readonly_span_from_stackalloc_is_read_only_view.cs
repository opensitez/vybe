// vybe-test: csharp/csharp_stackalloc_span/readonly_span_from_stackalloc_is_read_only_view
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.ReadOnlySpan<int> view=stackalloc int[2]{3,4}; __Check((view[1]).ToString(), "4");
