// vybe-test: csharp/csharp_stackalloc_span/stackalloc_int_single_element_initializer
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> buf=stackalloc int[1]{42}; __Check((buf[0]).ToString(), "42");
