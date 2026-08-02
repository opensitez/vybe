// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_reverse_in_place_changes_order
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> span=stackalloc int[3]{1,2,3}; span.Reverse(); __Check((span[0]).ToString(), "3"); __Check((span[2]).ToString(), "1");
