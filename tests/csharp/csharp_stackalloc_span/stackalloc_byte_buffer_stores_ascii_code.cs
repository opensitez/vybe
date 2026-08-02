// vybe-test: csharp/csharp_stackalloc_span/stackalloc_byte_buffer_stores_ascii_code
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<byte> buf=stackalloc byte[2]{65,66}; __Check((buf[0]).ToString(), "65");
