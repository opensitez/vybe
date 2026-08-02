// vybe-test: csharp/csharp_stackalloc_span/stackalloc_int_three_element_initializer_reads_middle
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> buf=stackalloc int[3]{10,20,30}; __Check((buf[1]).ToString(), "20");
