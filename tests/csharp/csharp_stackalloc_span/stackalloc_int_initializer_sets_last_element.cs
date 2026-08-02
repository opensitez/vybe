// vybe-test: csharp/csharp_stackalloc_span/stackalloc_int_initializer_sets_last_element
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> buf=stackalloc int[4]{1,2,3,4}; __Check((buf[3]).ToString(), "4");
