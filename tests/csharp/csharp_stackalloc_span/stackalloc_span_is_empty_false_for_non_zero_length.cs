// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_is_empty_false_for_non_zero_length
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> span=stackalloc int[1]{1}; __Check((span.IsEmpty).ToString(), "False");
