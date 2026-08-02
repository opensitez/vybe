// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_length_after_slice_one_from_three
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> span=stackalloc int[3]{5,6,7}; __Check((span.Slice(1).Length).ToString(), "2");
