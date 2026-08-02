// vybe-test: csharp/csharp_stackalloc_span/stackalloc_int_write_through_span_persists_in_buffer
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> buf=stackalloc int[2]{1,2}; buf[1]=99; __Check((buf[1]).ToString(), "99");
