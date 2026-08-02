// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_overlapping_copy_to_self_is_allowed
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> span=stackalloc int[3]{1,2,3}; span.CopyTo(span.Slice(1)); __Check((span[2]).ToString(), "2");
