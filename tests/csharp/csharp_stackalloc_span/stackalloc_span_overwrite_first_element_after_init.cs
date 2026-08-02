// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_overwrite_first_element_after_init
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> span=stackalloc int[3]{1,2,3}; span[0]=100; __Check((span[0]).ToString(), "100");
