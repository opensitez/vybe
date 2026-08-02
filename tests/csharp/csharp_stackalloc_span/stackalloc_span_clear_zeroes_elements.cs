// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_clear_zeroes_elements
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> span=stackalloc int[2]{5,6}; span.Clear(); __Check((span[0]).ToString(), "0"); __Check((span[1]).ToString(), "0");
